/**
 * RSVP mariage — API Apps Script
 *
 * GET  ?q=Prénom Nom            -> cherche l'invité, renvoie son foyer
 * GET  ?q=Prénom Nom&foyer=XXX  -> idem, quand plusieurs foyers correspondent
 * POST (text/plain, JSON)       -> enregistre les réponses du foyer
 *
 * Déploiement : Déployer > Nouveau déploiement > Application web
 *   Exécuter en tant que : moi / Qui a accès : Tout le monde
 */

const SHEET_INVITES = 'Invités';
const SHEET_REPONSES = 'Réponses';
const CATEGORIES_ENFANT = ['Enfant -3 ans', 'Enfant 3-12 ans', 'Enfant +12 ans'];
const MAX_ENFANTS = 10;
const MAX_LEN = 300;

// ---------- Utilitaires ----------

function normalize_(s) {
  return String(s || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function clean_(s) {
  return String(s == null ? '' : s).replace(/^[=+\-@]+/, '').trim().slice(0, MAX_LEN);
}

function ouiNon_(v) {
  return v === true || v === 'Oui' ? 'Oui' : 'Non';
}

/** Lit l'onglet Invités et renvoie la liste des personnes (colonnes repérées par leur en-tête). */
function readInvites_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_INVITES);
  const values = sh.getDataRange().getValues();
  const head = values[0].map(String);
  const col = name => {
    const i = head.indexOf(name);
    if (i < 0) throw new Error('Colonne introuvable dans Invités : ' + name);
    return i;
  };
  const c = {
    foyer: col('Foyer'), prenom: col('Prénom'), nom: col('Nom'),
    plusUn: col('+1 autorisé'), enfants: col('Enfants autorisés'), brunch: col('Invité au brunch'),
    categorie: head.indexOf('Catégorie'),
  };
  return values.slice(1)
    .filter(r => String(r[c.prenom]).trim() && String(r[c.foyer]).trim())
    .map(r => ({
      foyer: String(r[c.foyer]).trim(),
      prenom: String(r[c.prenom]).trim(),
      nom: String(r[c.nom]).trim(),
      plusUn: r[c.plusUn] === true,
      enfants: r[c.enfants] === true,
      brunch: r[c.brunch] === true,
      categorie: CATEGORIES_ENFANT.includes(r[c.categorie]) ? r[c.categorie] : 'Adulte',
    }));
}

/** Construit la vue d'un foyer, avec les réponses déjà enregistrées pour pré-remplir. */
function foyerView_(invites, foyer) {
  const membres = invites.filter(p => p.foyer === foyer);
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_REPONSES);
  const rep = sh.getDataRange().getValues().slice(1).filter(r => String(r[1]) === foyer);
  const toObj = r => ({
    type: r[2], prenom: r[3], nom: r[4], categorie: r[5],
    samedi: r[6] === 'Oui', dimanche: r[7] === 'Oui', regime: r[8], allergies: r[9],
  });
  const existing = rep.map(toObj);
  const message = (rep.find(r => r[10]) || [])[10] || '';
  return {
    foyer: foyer,
    plusUn: membres.some(m => m.plusUn),
    enfantsAutorises: membres.some(m => m.enfants),
    brunch: membres.some(m => m.brunch),
    dejaRepondu: rep.length > 0,
    membres: membres.map(m => {
      const e = existing.find(x => x.type === 'Invité' && x.prenom === m.prenom && x.nom === m.nom);
      return { prenom: m.prenom, nom: m.nom, categorie: m.categorie, reponse: e || null };
    }),
    accompagnant: existing.find(x => x.type === 'Accompagnant') || null,
    enfants: existing.filter(x => x.type === 'Enfant'),
    message: message,
  };
}

// ---------- Recherche ----------

function doGet(e) {
  try {
    const q = normalize_(e.parameter.q);
    const tokens = q.split(' ').filter(Boolean);
    if (tokens.length < 2) {
      return json_({ ok: false, erreur: 'Indique ton prénom et ton nom.' });
    }
    const invites = readInvites_();
    // Une personne correspond si chaque mot tapé est un mot de son prénom + nom.
    const matches = invites.filter(p => {
      const words = normalize_(p.prenom + ' ' + p.nom).split(' ');
      return tokens.every(t => words.includes(t));
    });
    let foyers = [...new Set(matches.map(m => m.foyer))];
    if (e.parameter.foyer) foyers = foyers.filter(f => f === e.parameter.foyer);

    if (foyers.length === 0) {
      return json_({ ok: false, erreur: "On ne te trouve pas dans la liste. Vérifie l'orthographe, ou contacte Tim ou Océane." });
    }
    if (foyers.length > 1) {
      return json_({
        ok: true, choix: foyers.map(f => ({
          foyer: f,
          membres: invites.filter(p => p.foyer === f).map(p => p.prenom + ' ' + p.nom).join(', '),
        })),
      });
    }
    return json_({ ok: true, foyer: foyerView_(invites, foyers[0]) });
  } catch (err) {
    return json_({ ok: false, erreur: 'Erreur serveur : ' + err.message });
  }
}

// ---------- Enregistrement ----------

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const data = JSON.parse(e.postData.contents);
    const invites = readInvites_();
    const membres = invites.filter(p => p.foyer === data.foyer);
    if (!membres.length) return json_({ ok: false, erreur: 'Foyer inconnu.' });

    // Sécurité minimale : il faut connaître le nom d'un membre du foyer.
    const qn = normalize_(data.q);
    const autorise = membres.some(p => {
      const words = normalize_(p.prenom + ' ' + p.nom).split(' ');
      const t = qn.split(' ').filter(Boolean);
      return t.length >= 2 && t.every(x => words.includes(x));
    });
    if (!autorise) return json_({ ok: false, erreur: 'Accès refusé.' });

    const brunch = membres.some(m => m.brunch);
    const plusUn = membres.some(m => m.plusUn);
    const enfantsOk = membres.some(m => m.enfants);
    const dim = v => (brunch ? ouiNon_(v) : '');
    const now = new Date();
    const message = clean_(data.message);
    const rows = [];

    // Membres listés : uniquement ceux du foyer, jamais d'ajout libre.
    (data.membres || []).forEach(m => {
      const p = membres.find(x => x.prenom === m.prenom && x.nom === m.nom);
      if (!p) return;
      rows.push([now, data.foyer, 'Invité', p.prenom, p.nom, p.categorie,
        ouiNon_(m.samedi), dim(m.dimanche), clean_(m.regime), clean_(m.allergies), '']);
    });

    // Accompagnant : seulement si +1 autorisé.
    const a = data.accompagnant;
    if (plusUn && a && clean_(a.prenom)) {
      rows.push([now, data.foyer, 'Accompagnant', clean_(a.prenom), clean_(a.nom), 'Adulte',
        ouiNon_(a.samedi), dim(a.dimanche), clean_(a.regime), clean_(a.allergies), '']);
    }

    // Enfants : seulement si autorisés, catégorie imposée.
    if (enfantsOk) {
      (data.enfants || []).slice(0, MAX_ENFANTS).forEach(k => {
        if (!clean_(k.prenom) || !CATEGORIES_ENFANT.includes(k.categorie)) return;
        rows.push([now, data.foyer, 'Enfant', clean_(k.prenom), clean_(k.nom), k.categorie,
          ouiNon_(k.samedi), dim(k.dimanche), clean_(k.regime), clean_(k.allergies), '']);
      });
    }

    if (!rows.length) return json_({ ok: false, erreur: 'Aucune réponse à enregistrer.' });
    rows[0][10] = message;

    // Remplace les anciennes réponses du foyer.
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_REPONSES);
    const all = sh.getDataRange().getValues();
    for (let i = all.length - 1; i >= 1; i--) {
      if (String(all[i][1]) === data.foyer) sh.deleteRow(i + 1);
    }
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, erreur: 'Erreur serveur : ' + err.message });
  } finally {
    lock.releaseLock();
  }
}
