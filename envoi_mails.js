function envoie_mail_kholleur_feuille_notation(civilite_nom_prenom_kholleur,id_feuille_notation,nom_fichier_releve_notes,salle)
{
  const nom = civilite_nom_prenom_kholleur["nom"];
  const prenom = civilite_nom_prenom_kholleur["prénom"];
  const le_kholleur = retourne_kholleur(nom,prenom);
  if (le_kholleur == -1)
    return;
  const mail_kholleur = le_kholleur["mail"];
  if (!isValidEmail(mail_kholleur))
  {
    msg_log = "Erreur mail kholleur invalide : " + mail_kholleur + " khôlleur : " + le_kholleur;
    ecrit_log(LOG_ERROR,"envoie_mail_kholleur_feuille_notation",msg_log);
    return;
  }
  const infos_kholle = nom_fichier_releve_notes.split(" - ");
  const discipline = infos_kholle[1];
  const classe = infos_kholle[2];
  const date_kholle = infos_kholle[3];
  const objet_mail = "Grille de khôlles du " + date_kholle + " en " + discipline + " classe : " + classe;
  const modele_corps_message_kholleur = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("modele_corps_message_kholleur_grille").getValue();
  const nom_prenom_kholleur = civilite_nom_prenom_kholleur["prénom"] + " " + civilite_nom_prenom_kholleur["nom"];
  let corps_du_mail = nom_prenom_kholleur; 
  //corps_du_mail += ",<BR><BR>"+modele_corps_message_kholleur + date_kholle + " - salle : " + salle;
  corps_du_mail += ",<BR>Nous sommes ravis de vous accueillir en prépa pour interroger nos élèves.";
  corps_du_mail += "<BR>Veuillez trouver ci-dessous un lien vers la grille des khôlles du " + date_kholle + ".";
  corps_du_mail += "<BR>Vous passerez dans la salle suivante : " + salle + ".";
  const fichier_releve_notes = SpreadsheetApp.openById(id_feuille_notation);
  corps_du_mail += "<BR><BR>" + "<a href='"+fichier_releve_notes.getUrl()+"'> lien vers la saisie des notes et commentaires</a>";
  corps_du_mail += "<BR><BR>Amicalement,<BR>L'équipe pédagogique";
  try
  {
    MailApp.sendEmail(mail_kholleur, objet_mail,'' ,{htmlBody: corps_du_mail});
    msg_log = "Mail envoyé au kholleur : grille de la khôlle " + nom_prenom_kholleur + " pour la khôlle du " + date_kholle + " en " + discipline + " classe : " + classe + " dans la salle " +  salle + " (mail : " + mail_kholleur + ")";
    ecrit_log(LOG_INFO,msg_log);
  }
  catch (err)
  {
    msg_log = "Erreur envoie mail kholleur : " + mail_kholleur + " " + err;
    ecrit_log(LOG_ERROR,"envoie_mail_kholleur_feuille_notation",msg_log);
  }
}

function envoie_mail_kholleur_saisie_notes(civilite,nom,prenom,id_feuille_notation,nom_fichier_releve_notes)
{
  const le_kholleur = retourne_kholleur(nom,prenom);
  if (le_kholleur == -1)
    return;
  const mail_kholleur = le_kholleur["mail"];
  if (!isValidEmail(mail_kholleur))
  {
    msg_log = "Erreur mail khôlleur invalide : " + mail_kholleur + " khôlleur : " + le_kholleur;
    ecrit_log(LOG_ERROR,"envoie_mail_kholleur_saisie_notes",msg_log);
    return;
  }
  const infos_kholle = nom_fichier_releve_notes.split(" - ");
  const discipline = infos_kholle[1];
  const classe = infos_kholle[2];
  const date_kholle = infos_kholle[3];
  const objet_mail = "Relevé de notes pour la khôlle du " + date_kholle + " en " + discipline + " classe : " + classe;
  const modele_corps_message_kholleur = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("modele_corps_message_kholleur_saisie_notes").getValue();
  const nom_prenom_kholleur = nom + " " + prenom;
  let corps_du_mail = nom_prenom_kholleur; 
  corps_du_mail += ",<BR><BR>"+modele_corps_message_kholleur + date_kholle ;
  const fichier_releve_notes = SpreadsheetApp.openById(id_feuille_notation);
  corps_du_mail += "<BR><BR><BR>" + "<a href='"+fichier_releve_notes.getUrl()+"'> lien vers la saisie des notes et commentaires</a>";
   corps_du_mail += "<BR><BR>Amicalement,<BR>L'équipe pédagogique";
  try
  {
      MailApp.sendEmail(mail_kholleur, objet_mail,'' ,{htmlBody: corps_du_mail});
      msg_log = "Mail envoyé au kholleur : relevé de notes et commentaires " + nom_prenom_kholleur + " pour la khôlle du " + date_kholle + " en " + discipline + " classe : " + classe + " (mail : " + mail_kholleur + ")";
      ecrit_log(LOG_INFO,msg_log);
  }
  catch (err)
  {
    msg_log = "Erreur envoie mail kholleur : " + mail_kholleur + " " + err;
    ecrit_log(LOG_ERROR,"envoie_mail_kholleur_saisie_notes",msg_log);
  }
}

function showSidebarMailRappelKholleursFeuilleNotation() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('envoie_mail_notes_commentaires_kholleurs').setTitle("Envoi Mailfeuille notation aux examinateurs");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
  return htmlOutput ;
}

function showSidebarMailRappelProfesseurFeuilleNotation() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('envoie_mail_notes_commentaires_professeur').setTitle("Envoi Mail feuille notation au professeur référent");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
  return htmlOutput ;
}

function envoyer_mail_rappel_kholleurs_feuille_notation()
{
   const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  if (!seance_deja_enregistree(nom_onglet_planning))
  {
    let msg = "Envoi mail rappel impossible : ce planning n'a pas été enregistré : <BR><BR>" + nom_onglet_planning;
    msg += "<BR><BR>Veuillez enregistrer le planning au préalable.";
    SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue(msg);
    showSideBarInfos("Envoi mail rappel feuille notation aux khôlleurs");
    return;
  }
  showSidebarMailRappelKholleursFeuilleNotation();
}


function envoyer_mail_rappel_professeur_controle_notation()
{
   const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  if (!seance_deja_enregistree(nom_onglet_planning))
  {
    let msg = "Envoi mail rappel impossible : ce planning n'a pas été enregistré : <BR><BR>" + nom_onglet_planning;
    msg += "<BR><BR>Veuillez enregistrer le planning au préalable.";
    SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue(msg);
    showSideBarInfos("Envoi mail rappel feuille contrôle notes au professeur référent");
    return;
  }
  showSidebarMailRappelProfesseurFeuilleNotation();
}

function envoie_mail_kholleur_confirmation_disponibilite(civilite,prenom,nom,date_kholle,discipline,classe,heure_debut,heure_fin)
{
  const le_kholleur = retourne_kholleur(nom,prenom);
  const plage_horaire = " de " + heure_debut + " à " + heure_fin;
  if (le_kholleur == -1)
    return;
  const mail_kholleur = le_kholleur["mail"];
  if (!isValidEmail(mail_kholleur))
  {
    msg_log = "Erreur mail eleve invalide : " + mail_kholleur + " khôlleur : " + le_kholleur;
    ecrit_log(LOG_ERROR,"envoie_mail_kholleur_confirmation_disponibilite",msg_log);
    return;
  }
  const objet_mail = "Confirmation de disponibilité pour la khôlle du " + date_kholle + plage_horaire + " en " + discipline + " classe : " + classe;
  const modele_confirmation_dispo_kholleur = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("modele_confirmation_dispo_kholleur").getValue();
  const nom_prenom_kholleur = prenom + " " + nom;
  let corps_du_mail = nom_prenom_kholleur; 
  corps_du_mail += ",<BR><BR>"+modele_confirmation_dispo_kholleur + " " + date_kholle + plage_horaire;
  let form_infos = cree_formulaire_confirmation_dispo_kholleur(civilite,prenom,nom,date_kholle,discipline,classe,plage_horaire);
  const form_link = form_infos["url"];
  const form_id = form_infos["id"];
  const sh_confirmations_dispo_kholleurs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CONFIRMATION_DISPO_KHOLLEURS);
  sh_confirmations_dispo_kholleurs.appendRow([civilite,nom,prenom,classe,discipline,date_kholle,plage_horaire,"","",new Date(),form_id]);
  corps_du_mail += "<BR><BR>" + "<a href='"+form_link+"'> lien vers la demande de confirmation de disponibilité</a>";
   corps_du_mail += "<BR><BR>Amicalement,<BR>L'équipe pédagogique";
  try
  {
    MailApp.sendEmail(mail_kholleur, objet_mail,'' ,{htmlBody: corps_du_mail});
    msg_log = "Confirmation disponibilité : Mail envoyé au kholleur " + nom_prenom_kholleur + " pour la khôlle du " + date_kholle + plage_horaire + " en " + discipline + " classe : " + classe + " (mail : " + mail_kholleur + ")";
    ecrit_log(LOG_INFO,msg_log);
  }
  catch (err)
  {
    msg_log = "Erreur envoie mail kholleur : " + mail_kholleur + " " + err;
    ecrit_log(LOG_ERROR,"envoie_mail_kholleur_confirmation_disponibilite",msg_log);
  }
}

function envoie_mail_professeur_controle_notation(le_professeur_referent,id_feuille_controle_professeur,nom_fichier_controle_professeur)
{
  let le_prof_split = le_professeur_referent.split(" ");
  const nom_professeur = le_prof_split[1];
  const prenom_professeur = le_prof_split[2];
  const le_professeur = retourne_professeur(nom_professeur,prenom_professeur);
  if (le_professeur == -1)
    return;
  const mail_professeur = le_professeur["mail"];
  const infos_kholle = nom_fichier_controle_professeur.split(" - ");
  const discipline = infos_kholle[1];
  const classe = infos_kholle[2];
  const date_kholle = infos_kholle[3];
  const objet_mail = "Vérification des notes et commentaires des examinateurs pour la khôlle du " + date_kholle + " en " + discipline + " classe : " + classe;
  const modele_corps_message_professeur = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("modele_corps_message_professeur").getValue();
  const nom_prenom_professeur = prenom_professeur + " " + nom_professeur;
  let corps_du_mail = nom_prenom_professeur; 
  corps_du_mail += ",<BR>Cher Professeur(e)";
  corps_du_mail += ",<BR><BR>Veuillez trouver ci-dessous le lien permettant de valider les notes et les commentaires de la khôlle du " + date_kholle + ".";
  const fichier_releve_notes = SpreadsheetApp.openById(id_feuille_controle_professeur);
  corps_du_mail += "<BR><BR>" + "<a href='"+fichier_releve_notes.getUrl()+"'> lien vers le fichier de contrôle des notes et commentaires</a>";
  corps_du_mail += "<BR><BR>Une fois validés par vous, les commentaires seront automatiquement envoyés aux élèves.";
  corps_du_mail += "<BR>Il ne vous restera plus qu’à copier-coller ces notes sur École Directe.";
  corps_du_mail += "<BR>Merci beaucoup.";
  corps_du_mail += "<BR><BR>Amicalement,<BR>L'équipe pédagogique";
  try
  {
  MailApp.sendEmail(mail_professeur, objet_mail,'' ,{htmlBody: corps_du_mail});
    msg_log = "Mail envoyé au professeur " + nom_prenom_professeur + " pour la khôlle du " + date_kholle + " en " + discipline + " classe : " + classe + " (mail : " + mail_professeur + ")";
    ecrit_log(LOG_INFO,msg_log);
  }
  catch (err)
  {
    msg_log = "Erreur envoie mail professeur : " + mail_professeur + " " + err;
    ecrit_log(LOG_ERROR,"envoie_mail_professeur_controle_notation",msg_log);
  }
}

function envoie_mail_eleve_convocation_kholle(civilite_nom_prenom_kholleur,eleve,mail,nom_fichier_releve_notes,salle)
{ 
  const infos_kholle = nom_fichier_releve_notes.split(" - ");
  const discipline = infos_kholle[1];
  const classe = infos_kholle[2];
  const date_kholle = infos_kholle[3];
  const heure_preparation = eleve["heure_preparation"];
  const heure_passage = eleve["heure_passage"];
  const civilite_nom_kholleur = civilite_nom_prenom_kholleur["civilité"] + " " + civilite_nom_prenom_kholleur["nom"];
  const objet_mail = "Convocation pour la khôlle du " + date_kholle + " à " +  heure_preparation + " en " + discipline + " avec " + civilite_nom_kholleur; 
  const modele_corps_message_eleve = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("modele_corps_message_etudiant_convocation").getValue();
  const prenom_nom_eleve = eleve["prénom"] + " " + eleve["nom"];
  let corps_du_mail = prenom_nom_eleve;
  corps_du_mail += ",<BR><BR>";
  corps_du_mail += modele_corps_message_eleve + date_kholle + " :";
  corps_du_mail += "<BR><BR>";
  corps_du_mail += "<ul>";
  corps_du_mail += "<li><p>en " + discipline + "</p></li>";
  corps_du_mail += "<li><p>avec " + civilite_nom_kholleur + "</p></li>";
  corps_du_mail += "<li><p>de " + replaceAll(heure_passage,"-","à") + "</p></li>";
  corps_du_mail += "<li><p>début préparation à " + heure_preparation + "</p></li>";
  corps_du_mail += "<li><p>dans la salle : " + classe + "</p></li>";
  corps_du_mail += "</ul>";
  corps_du_mail += "<BR><BR>Bien cordialement,<BR>L'équipe pédagogique";
  try
  {
    MailApp.sendEmail(mail, objet_mail,'' ,{htmlBody: corps_du_mail});
    msg_log = "Mail envoyé à l'étudiant " + prenom_nom_eleve + " pour la convocation à la khôlle du " + date_kholle + " à " + heure_preparation + " en " + discipline + " classe : " + classe + " avec " + civilite_nom_prenom_kholleur["civilité"] + " " + civilite_nom_prenom_kholleur["prénom"] + " " + civilite_nom_prenom_kholleur["nom"];
    msg_log +=  " dans la salle " + salle;
    msg_log += " (mail : " + mail + ")"
    ecrit_log(LOG_INFO,msg_log);
  }
  catch (err)
  {
    msg_log = "Erreur envoie mail étudiant : " + mail + " " + err;
    ecrit_log(LOG_ERROR,"envoie_mail_eleve_convocation_kholle",msg_log);
  }
}

function envoie_mail_eleve_resultat_kholle(kholleur,eleve,discipline,classe,date_kholle,note,commentaire)
{
  const kholleur_split = kholleur.split(" ");
  const civilite_kholleur = kholleur_split[0];
  const nom_kholleur = kholleur_split[1];
  const civilite_nom_kholleur = civilite_kholleur + " " + nom_kholleur;
  const objet_mail = "Résultats pour la khôlle du " + date_kholle  + " en " + discipline + " avec " + civilite_nom_kholleur; 
  const modele_corps_message_eleve = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("modele_corps_message_etudiant_resultats").getValue();
  const prenom_nom_eleve = eleve["prenom"] + " " + eleve["nom"];
  const mail = eleve["mail"];
  if (!isValidEmail(mail))
  {
    msg_log = "Erreur mail eleve invalide : " + mail + " eleve : " + prenom_nom_eleve;
    ecrit_log(LOG_ERROR,"envoie_mail_eleve_resultat_kholle",msg_log);
    return;
  }
  let corps_du_mail = prenom_nom_eleve;
  corps_du_mail += ",<BR><BR>";
  corps_du_mail += modele_corps_message_eleve + " " + date_kholle;
  corps_du_mail += " avec comme examinateur : " + civilite_nom_kholleur + " en " + discipline + ".";
  const parametre_envoi_resultats_kholle = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("parametre_envoi_resultats_etudiants").getValue();
  if (parametre_envoi_resultats_kholle == ENVOI_NOTE_ET_COMMENTAIRE)
  {
    corps_du_mail += "<br><br>Votre note : " + note + "/20";
    corps_du_mail += "<br>Commentaire de " + kholleur + " :<br>" + commentaire;
  }
  else
    corps_du_mail += " <br><br>Commentaire de " + kholleur + " :<br>" + commentaire;
 
  corps_du_mail += "<BR><BR>Bien cordialement,<BR>L'équipe pédagogique";
  try
  {
    MailApp.sendEmail(mail, objet_mail,'' ,{htmlBody: corps_du_mail});
    msg_log = "Mail envoyé à l'étudiant " + prenom_nom_eleve + " pour les résultats la khôlle du " + date_kholle + " en " + discipline + " classe : " + classe + " avec l'examinateur : " + civilite_nom_kholleur;
    msg_log += " (mail : " + mail + ")";
    if (parametre_envoi_resultats_kholle == ENVOI_NOTE_ET_COMMENTAIRE)
    {
      msg_log += "\n Note : " + note + "/20";
      msg_log += "\nCommentaire de " + kholleur + " : " + commentaire;
    }
    else
      corps_du_mail += "\nCommentaire de " + kholleur + " : " + commentaire;
    ecrit_log(LOG_INFO,msg_log);
    }
     catch (err)
  {
    msg_log = "Erreur envoie mail étudiant : " + mail + " " + err;
    ecrit_log(LOG_ERROR,"envoie_mail_eleve_resultat_kholle",msg_log);
  }
  
}

function test_envoie_mails_rappel_feuille_notation_kholleurs()
{
  envoie_mails_rappel_feuille_notation_kholleurs([1,2]);
}

function envoie_mails_rappel_feuille_notation_kholleurs(les_indices_kholleurs)
{
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  const infos_planning = retourne_infos_from_nom_onglet_planning(nom_onglet_planning);
  const discipline = infos_planning["discipline"];
  const annee = infos_planning["annee"];
  let date_kholle = infos_planning["date_kholle"];
  const sh_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  let kholleurs_contactes = [];
  const nb_examinateurs = sh_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
  const donnees_kholleurs = lit_donnees_kholleurs_feuille_planning(sh_planning,nb_examinateurs);
  for (let ind_kholleur of les_indices_kholleurs)
  {
    ind_kholleur = parseInt(ind_kholleur);
    const kholleur = donnees_kholleurs[ind_kholleur-1][0];
     const kholleur_split = kholleur.split(" ");
    const civilite_kholleur = kholleur_split[0];
    const nom_kholleur = kholleur_split[1];
    const prenom_kholleur = kholleur_split[2];
    const row = sh_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES+ind_kholleur).getValue();
    const salle = sh_planning.getRange(row,1).getValue();
    const civilite_nom_prenom_kholleur = {"civilité":civilite_kholleur,"nom":nom_kholleur,"prénom":prenom_kholleur};
    const nom_fichier_releve_notes = "Relevé examinateur " + kholleur + " - " + discipline + " - " + annee + " - " + date_kholle;
    const id_feuille_notation = sh_planning.getRange(LIGNE_MEMO+2+ind_kholleur,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
    if (id_feuille_notation == "")
    {
      msg_log = "Erreur envoi mail rappel feuille de notation aux examinateurs : id_feuille_notation vide" + nom_fichier_releve_notes;
      ecrit_log(LOG_ERROR,"envoie_mails_rappel_feuille_notation_kholleurs",msg_log);
    }
    if (validateFileId(id_feuille_notation))
    {
      kholleurs_contactes.push(kholleur);
      envoie_mail_kholleur_feuille_notation(civilite_nom_prenom_kholleur,id_feuille_notation,nom_fichier_releve_notes,salle);
      let msg = "Les examinateurs suivants ont été contactés par mail avec le lien du fichier de notation des étudiants pour la khôlle :<BR><H2>" + annee + " " + discipline + " " + date_kholle + "</H2> : <BR>";
      for (const kholleur of kholleurs_contactes)
        msg += "<BR>" + kholleur;  
      SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue(msg);
      showSideBarInfos("Envoi mail rappel feuille notation aux khôlleurs");
    }
    else
    {
      msg_log = "Erreur envoi mail rappel feuille de notation aux examinateurs : id_feuille_notation format incorrect" + id_feuille_notation + " - " + nom_fichier_releve_notes;
      ecrit_log(LOG_ERROR,"envoie_mails_rappel_feuille_notation_kholleurs",msg_log);
    }
  }
}

function envoie_mail_rappel_professeur()
{
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  const infos_planning = retourne_infos_from_nom_onglet_planning(nom_onglet_planning);
  const discipline = infos_planning["discipline"];
  const annee = infos_planning["annee"];
  let date_kholle = infos_planning["date_kholle"];
  const sh_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  const le_professeur_referent = sh_planning.getRange("H1").getValue();
  const nom_fichier_controle_professeur = "Relevé professeur " + le_professeur_referent + " - " + discipline + " - " + annee + " - " + date_kholle;
  const id_feuille_controle_professeur = sh_planning.getRange(LIGNE_MEMO+2,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
  if (id_feuille_controle_professeur == "")
  {
    msg_log = "Erreur envoi mail rappel feuille de notation aux examinateurs : id_feuille_notation vide" + nom_fichier_controle_professeur;
    ecrit_log(LOG_ERROR,"envoie_mail_rappel_professeur",msg_log);
  }
  if (validateFileId(id_feuille_controle_professeur))
  {
    envoie_mail_professeur_controle_notation(le_professeur_referent,id_feuille_controle_professeur,nom_fichier_controle_professeur);
    let msg = "Mail de rappel avec le lien pour les notes / commentaires de la khôlle :<BR><H2>" + annee + " " + discipline + " " + date_kholle + "</H2> bien envoyé au professeur : <BR><BR>" + le_professeur_referent;
      SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue(msg);
      showSideBarInfos("Envoi mail de rappel du lien de contrôle des résultats au professeur référent");
  }
  else
  {
    msg_log = "Erreur envoi mail rappel feuille de notation au professeur : format id_feuille_controle_professeur incorrect " + id_feuille_controle_professeur + " - " + nom_fichier_controle_professeur;
    ecrit_log(LOG_ERROR,"envoie_mail_rappel_professeur",msg_log);
  }
}