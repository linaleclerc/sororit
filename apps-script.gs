function doPost(e) {
  var p = e.parameter;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Réponses au sondage -> onglet "Sondage" (pas de mail)
  if (p.type === 'survey') {
    var sh = ss.getSheetByName('Sondage') || ss.insertSheet('Sondage');
    if (sh.getLastRow() === 0) sh.appendRow(['Date', 'Intérêt', 'Fonctionnalité', 'Profil', 'Recommandation', 'Commentaire']);
    sh.appendRow([new Date(), p.q_interet || '', p.q_fonctionnalite || '', p.q_profil || '', p.q_recommandation || '', p.q_commentaire || '']);
    return ContentService.createTextOutput('ok');
  }

  // Inscriptions -> premier onglet + mail de confirmation
  var name = p.name || '', email = p.email || '', msg = p.message || '', fr = p.lang === 'fr';
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return ContentService.createTextOutput('invalid');

  ss.getSheets()[0].appendRow([new Date(), name, email, msg]);

  MailApp.sendEmail({
    to: email,
    name: 'SororIT',
    subject: fr ? 'Ton inscription à SororIT est confirmée 💜' : "You're on the SororIT list 💜",
    body: fr
      ? 'Coucou ' + name + ',\n\nMerci ! Ton inscription à SororIT est bien enregistrée. Tu seras prévenue dès que la communauté ouvre.\n\nÀ très vite 💜\nL\'équipe SororIT'
      : 'Hi ' + name + ',\n\nThank you! Your SororIT registration is confirmed. We\'ll let you know the moment the community goes live.\n\nSee you soon 💜\nThe SororIT team'
  });
  return ContentService.createTextOutput('ok');
}
