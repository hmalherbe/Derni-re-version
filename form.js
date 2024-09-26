function cree_formulaire_confirmation_dispo_kholleur(civilite,prenom,nom,date_kholle,discipline,classe,plage_horaire) 
{
  let titre = 'Demande de confirmation de disponibilité pour une Khôlle de Prépa D1 à Stanislas Nice';
  let form = FormApp.create(titre);
  let item = form.addMultipleChoiceItem();
  form.setDescription("Discipline : " + discipline + "\n Date de la khôlle " + date_kholle + " " + plage_horaire + "\nen " + classe);
  item.setTitle('Confirmez-vous votre disponibilité pour la kholle décrite ci-dessus ?');
  item.setChoices([
    item.createChoice('Oui, je confirme ma disponibilité'),
    item.createChoice('Non, je ne suis pas disponible'),
  ]);
  const annee_scolaire = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("Annee_scolaire").getValue();
  const le_kholleur = civilite + " " + nom + " " + prenom;
  const chemin_formulaire_confirmation_dispo_kholleur = annee_scolaire + "/" + "Notes/" + discipline + "/Examinateurs/" + le_kholleur;
  supprime_si_existe_formulaire(civilite,prenom,nom,date_kholle,discipline,classe);
  const rootFolderId = DriveApp.getRootFolder().getId();
  const folder_id = getFolderIdByPath(rootFolderId,chemin_formulaire_confirmation_dispo_kholleur);
  if (folder_id==null)
    return null;
  form.setTitle(titre);
  const new_form_id = moveFormToFolder(form.getId(),folder_id);
  form = FormApp.openById(new_form_id);
  return {"id":form.getId(),"url":form.getPublishedUrl()};
}

  
