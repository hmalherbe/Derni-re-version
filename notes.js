function envoyer_notes_commentaires_eleves()
{
    const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
    const id_courant = SpreadsheetApp.getActiveSpreadsheet().getId();
    supprime_tous_releves_notes();
    const sh_nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
    const id_feuille_notation = sh_nom_onglet_planning.getRange(LIGNE_MEMO+2,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
    if (id_feuille_notation != "")
    {
      try
      {
        const feuille_releve_notes = SpreadsheetApp.openById(id_feuille_notation).getSheets()[0].getName();
        SpreadsheetApp.openById(id_courant);
        copySheetToAnotherSpreadsheet(id_feuille_notation,SpreadsheetApp.getActiveSpreadsheet().getId());
        SpreadsheetApp.openById(id_courant);
        const sh_notes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Copie de " + feuille_releve_notes).activate();
        sh_notes.setName(feuille_releve_notes);
        sh_notes.getRange("E6:F6").clearContent();
        sh_notes.getRange("E6:F6").clearDataValidations();
        const sh_temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CALCUL_TEMPORAIRES);
        sh_temp.getRange(2,1).setValue(feuille_releve_notes);
      }
      catch (err)
      {
        msg_log = "Erreur feuille notation professeur : " + id_feuille_notation;
        ecrit_log(LOG_ERROR,"envoyer_notes_commentaires_eleves",msg_log);
        return;
      }
      showSidebarNotesEleves();
    }
}

function supprime_tous_releves_notes()
{
  for (const feuille of SpreadsheetApp.getActiveSpreadsheet().getSheets())
  {
    if (feuille.getName().startsWith("Relevé professeur"))
    {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(feuille.getName());
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
    }
  }
}

function showSidebarNotesEleves() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('envoie_notes_commentaires_eleves').setTitle("Envoi des résultats de la khôlle aux étudiants");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
  return htmlOutput ;
}

function construit_liste_eleves() {
  let html = "";
  const sh_temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CALCUL_TEMPORAIRES);
  const sheet_name = sh_temp.getRange(2,1).getValue();
  const feuille_notes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).activate();
  const donnees_notes = feuille_notes.getDataRange().getValues().slice(2);
  const feuille_notes_split = feuille_notes.getName().split(" - ");
  const discipline = feuille_notes_split[1];
  const classe = feuille_notes_split[2];
  const date_kholle = feuille_notes_split[3];
  let css = "";
  html = "<H1>" + discipline + " " + classe + " " + date_kholle + "</H1>";
  html += "<div>";
  html += "<input type='checkbox' id ='chk_eleves_tous' value='-1' checked onchange='change_eleves_tous(this);'>Envoyer à tous les étudiants</label><br><br>";
  html += "</div>";
  html += "<div id='div_chk_eleves'  style='visibility:hidden;'>";
  for (let i = 0;i < donnees_notes.length;i++)
  {
    const eleve =  donnees_notes[i][0];
    html += "<input type='checkbox' name='chk_eleves' value='"+i+"' onchange='change_chk_eleves(this);'>" + eleve +"</label><br>";
  }
  html += "</div>";
  return html;
}

function construit_liste_kholleurs() {
  let html = "";
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  const infos_planning = retourne_infos_from_nom_onglet_planning(nom_onglet_planning);
  const discipline = infos_planning["discipline"];
  const annee = infos_planning["annee"];
  let date_kholle = infos_planning["date_kholle"];
  /*const date_split = date_kholle.split("/");
  date_kholle = new Date(date_split[2],date_split[1]-1,date_split[0]);*/
  const sh_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  const nb_examinateurs = sh_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
  html = "<H1>" + discipline + " " + annee + " " + date_kholle + "</H1>";
  html += "<div>";
  const donnees_kholleurs = lit_donnees_kholleurs_feuille_planning(sh_planning,nb_examinateurs);
  for (let i = 1; i <= nb_examinateurs;i++)
  {
    const kholleur = donnees_kholleurs[i-1][0];
    html += "<input type='checkbox' name='chk_kholleurs' value='"+i+"' onchange='change_chk_kholleurs(this);'>" + kholleur +"</label><br>";
  }
  html += "</div><BR><BR>";
  return html;
}

function construit_message_professeur() {
  let html = "";
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  const infos_planning = retourne_infos_from_nom_onglet_planning(nom_onglet_planning);
  const discipline = infos_planning["discipline"];
  const annee = infos_planning["annee"];
  let date_kholle = infos_planning["date_kholle"];
  /*const date_split = date_kholle.split("/");
  date_kholle = new Date(date_split[2],date_split[1]-1,date_split[0]);*/
  const sh_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  html = "<H1>" + discipline + " " + annee + " " + date_kholle + "</H1>";
  html += "<div>";
  const le_professeur_referent = sh_planning.getRange("H1").getValue();
  html += "Confirmez-vous l'envoi du mail de rappel avec le lien vers la feuille de notes / commentaire au professeur référent : <BR><BR>" + le_professeur_referent + " ?";
  html += "</div><BR><BR>";
  return html;
}

function test_envoie_resultats_eleves()
{
  const les_eleves = {"tous":true,"liste_eleves":[]};
  envoie_resultats_eleves(les_eleves);
}

function envoie_resultats_eleves(les_eleves)
{
  const sh_temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CALCUL_TEMPORAIRES);
  const sheet_name = sh_temp.getRange(2,1).getValue();
  const sheet_name_split = sheet_name.split(" - ");
  const discipline = sheet_name_split[1];
  const classe = sheet_name_split[2];
  const date_kholle = sheet_name_split[3];
  let msg = "Envoi des résultats aux élèves par mail en cours .... <BR><BR>Pour la khôlle : " + discipline + " " + classe + " " + date_kholle;
  msg += "...<BR><BR> Veuillez patienter ...";
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue(msg);
  showSideBarInfos("Envoi des résultats aux étudiants");
  const feuille_notes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  const donnees_notes = feuille_notes.getDataRange().getValues().slice(2);
  const parametre_envoi_resultats_etudiants = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("parametre_envoi_resultats_etudiants").getValue();
  let les_indices_eleves = [];
  if (les_eleves["tous"])
  {
    for (let i = 0;i < donnees_notes.length;i++)
      les_indices_eleves.push(i);
  }
  else
  {
    les_indices_eleves = [...les_eleves["liste_eleves"]];
  }
  for (const ind of les_indices_eleves)
  {
    const nom_prenom_eleve = donnees_notes[ind][0];
    const nom_prenom_eleve_split = nom_prenom_eleve.split(" ");
    const nom_eleve = nom_prenom_eleve_split[0];
    const prenom_eleve = nom_prenom_eleve_split[1];
    const kholleur = donnees_notes[ind][1];
    const l_eleve = retourne_eleve(nom_eleve,prenom_eleve);
    if (l_eleve==-1)
      return;
    const note = donnees_notes[ind][2];
    const commentaire = donnees_notes[ind][3];
    let envoi_mail = true;
    if (note == "" && parametre_envoi_resultats_etudiants == ENVOI_NOTE_ET_COMMENTAIRE)
    {
      msg_log = "Note non saisie pour l'étudiant : " + nom_eleve + " " + prenom_eleve + " kholle : " + date_kholle + " discipline : " + discipline + " classe : " + classe + " khôlleur : " + kholleur;
      ecrit_log(LOG_WARNING,"envoie_resultats_eleves",msg_log); 
      envoi_mail = false;
    }
    if (commentaire == "" && parametre_envoi_resultats_etudiants == ENVOI_NOTE_ET_COMMENTAIRE)
    {
      msg_log = "commentaire non saisi pour l'étudiant : " + nom_eleve + " " + prenom_eleve + " kholle : " + date_kholle + " discipline : " + discipline + " classe : " + classe + " khôlleur : " + kholleur;
      ecrit_log(LOG_WARNING,"envoie_resultats_eleves",msg_log); 
      envoi_mail = false;
    }
    if (envoi_mail && note.toUpperCase() != "ABS")
      envoie_mail_eleve_resultat_kholle(kholleur,l_eleve,discipline,classe,date_kholle,note,commentaire);
  }
  const nb_mails_envoyes = les_indices_eleves.length + "/" + donnees_notes.length;
  let sh_plannings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_PLANNING_KHOLLES);
  let i = 2;
  while (sh_plannings.getRange(i,1).getValue() != "")
  {
    let annee_pl = sh_plannings.getRange(i,1).getValue();
    if (classe == annee_pl)
    {
      let date_kholle_pl = formatDate(sh_plannings.getRange(i,2).getValue());
      if (date_kholle == date_kholle_pl)
      {
        let discipline_pl = sh_plannings.getRange(i,3).getValue();
        if (discipline == discipline_pl)
        {
          sh_plannings.getRange(i,7).setValue("Oui : " + nb_mails_envoyes);
          sh_plannings.activate();
          let msg_termine = "Résultats envoyés par mail aux étudiants <BR> <BR>";
          msg_termine += nb_mails_envoyes + " mails envoyés";
          SpreadsheetApp.getActiveSpreadsheet().getRangeByName
          ("message_sidebar").setValue(msg_termine);
          showSideBarInfos("Envoi des résultats aux étudiants");
          return;
        }
      }
    }
    i++;
  }
}
