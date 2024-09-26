function maj_notes_commentaires_kholleurs()
{
  const feuille_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_PLANNING_KHOLLES);
  const donnees_planning = feuille_planning.getDataRange().getValues();
  const textStyle_valide = SpreadsheetApp.newTextStyle()
  .setForegroundColor(COULEUR_NOTES_ET_COMMENTAIRES_KHOLLEURS_VALIDEES)
  .build();
  const textStyle_non_valide = SpreadsheetApp.newTextStyle()
  .setForegroundColor(COULEUR_NOTES_ET_COMMENTAIRES_KHOLLEURS_NON_VALIDEES)
  .build();  
  for (let i = 0;i < donnees_planning.length;i++)
  {
    if (donnees_planning[i][3] == "Oui")
    {
      const classe = donnees_planning[i][0];
      const date_kholle =  formatDate(donnees_planning[i][1]);
      const discipline =  donnees_planning[i][2];
      const nom_onglet_planning = construit_nom_onglet_planning(classe,date_kholle,discipline);
      const sh_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
      if (sh_planning==null)
        continue;
      const nb_examinateurs = sh_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
      let cr_notes_commentaires = "";
      let kholleurs_pos_saisie_complete = [];
      let pos_debut = 0;
      let pos_fin = 0;
      let au_moins_une_saisie_incomplete = false;
      let nb_validations = 0;
      for (let j = 1 ; j <= nb_examinateurs;j++)
      {
         const id_nom_fichier_releve_notes = sh_planning.getRange(LIGNE_MEMO+2+j,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
         if (id_nom_fichier_releve_notes!="")
         {
            const row = sh_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES+j).getValue();
            const kholleur = sh_planning.getRange(row,2).getValue();
            let feuille_releve_notes = SpreadsheetApp.openById(id_nom_fichier_releve_notes);
            let releve_notes = feuille_releve_notes.getDataRange().getValues().slice(2);
            releve_notes = releve_notes.filter(row => row[0] != "");
            let notes_remplies = releve_notes.filter(function(row) 
            {
              return Number.isInteger(row[2]) || row[2].toString().toUpperCase() == "ABS";
            });
            let commentaires_remplis = releve_notes.filter(function(row) 
            {
              return row[3] != "";
            });
            if (cr_notes_commentaires != "")
            {
              pos_fin += 1;
              pos_debut = pos_fin;
              cr_notes_commentaires += "\n";
            }
            let notes_commentaires_kholleur = kholleur + " : " + notes_remplies.length + "/" + releve_notes.length + " notes - " + commentaires_remplis.length + "/" + releve_notes.length + " commentaires";
            cr_notes_commentaires += notes_commentaires_kholleur; 
            pos_fin += notes_commentaires_kholleur.length;
            insere_case_a_cocher_validation(feuille_releve_notes);
            if (notes_remplies.length == releve_notes.length && commentaires_remplis.length == releve_notes.length) 
            {
              const etat_validation = feuille_releve_notes.getRange('E6').getValue();
              if (etat_validation)
              {
                 cr_notes_commentaires += " - relevé validé"; 
                pos_fin += 16;
                nb_validations += 1;
              }
              kholleurs_pos_saisie_complete.push([pos_debut,pos_fin,etat_validation]);
            }
            else
            {
              au_moins_une_saisie_incomplete = true;
              feuille_releve_notes.getRange('E6').setValue(false);
             }
          }
        }
        const richText = SpreadsheetApp.newRichTextValue().setText(cr_notes_commentaires); 
        if (au_moins_une_saisie_incomplete || cr_notes_commentaires == "")
        {
          feuille_planning.getRange(i+1,5).setBackground(null);
          for (let j = 0; j < kholleurs_pos_saisie_complete.length ; j++)
          {
            if (kholleurs_pos_saisie_complete[j][2])
              richText.setTextStyle(kholleurs_pos_saisie_complete[j][0], kholleurs_pos_saisie_complete[j][1],textStyle_valide);
            else
              richText.setTextStyle(kholleurs_pos_saisie_complete[j][0], kholleurs_pos_saisie_complete[j][1],textStyle_non_valide);
          } 
        }
        else
        {
          if (nb_validations == nb_examinateurs)
            feuille_planning.getRange(i+1,5).setBackground(COULEUR_NOTES_ET_COMMENTAIRES_KHOLLEURS_VALIDEES);
          else
            feuille_planning.getRange(i+1,5).setBackground(COULEUR_NOTES_ET_COMMENTAIRES_KHOLLEURS_NON_VALIDEES);
        }
        feuille_planning.getRange(i+1,5).setRichTextValue(richText.build());     
      }
    }
    //feuille_planning.autoResizeColumns(1,7);
}

function maj_notes_commentaires_professeurs()
{
const feuille_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_PLANNING_KHOLLES);
  const donnees_planning = feuille_planning.getDataRange().getValues();
  for (let i = 0;i < donnees_planning.length;i++)
  {
    if (donnees_planning[i][3] == "Oui")
    {
      const classe = donnees_planning[i][0];
      const date_kholle =  formatDate(donnees_planning[i][1]);
      const discipline =  donnees_planning[i][2];
      const nom_onglet_planning = construit_nom_onglet_planning(classe,date_kholle,discipline);
      const sh_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
      if (sh_planning==null)
        continue;
      const le_professeur_referent = sh_planning.getRange("H1").getValue();
      const id_nom_fichier_controle_professeur = sh_planning.getRange(LIGNE_MEMO+2,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
      if (id_nom_fichier_controle_professeur!="")
      {
        const fichier_controle_professeur = SpreadsheetApp.openById(id_nom_fichier_controle_professeur);
        let donnees_fichier_controle_professeur = fichier_controle_professeur.getDataRange().getValues();
        if (fichier_controle_professeur != null)
        {
          const nb_examinateurs = sh_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
          
          let etat_validation_kholleurs = [];
          for (let j = 1 ; j <= nb_examinateurs;j++)
          {
            const id_nom_fichier_releve_notes = sh_planning.getRange(LIGNE_MEMO+2+j,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
            if (id_nom_fichier_releve_notes!="")
            {
              const row = sh_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES+j).getValue();
              const kholleur = sh_planning.getRange(row,2).getValue();
              let feuille_releve_notes = SpreadsheetApp.openById(id_nom_fichier_releve_notes);
              let releve_notes = feuille_releve_notes.getDataRange().getValues().slice(2);
              releve_notes = releve_notes.filter(row => row[0] != "");
              let notes_remplies = releve_notes.filter(function(row) 
              {
                return Number.isInteger(row[2]) || row[2].toString().toUpperCase() == "ABS";
              });
              let commentaires_remplis = releve_notes.filter(function(row) 
              {
                return row[3] != "";
              });
              for (const note of notes_remplies)
              {
                const nom_prenom_eleve = note[0];
                let trouve = false;
                for (let j = 0;j < donnees_fichier_controle_professeur.length && !trouve ; j++)
                {
                  if (donnees_fichier_controle_professeur[j][0] == nom_prenom_eleve)
                  {
                    trouve = true;
                    donnees_fichier_controle_professeur[j][2] = note[2];
                  }
                }
                if (!trouve)
                {
                  msg_log = "Erreur eleve non trouvé : " + nom_prenom_eleve + " dans le fichier de contrôle des notes du professeur ";
                  msg_log += le_professeur_referent + " pour la kholle : " + feuille_releve_notes.getName();
                  ecrit_log(LOG_ERROR,msg_log);
                }
              }
              for (const commentaire of commentaires_remplis)
              {
                const nom_prenom_eleve = commentaire[0];
                let trouve = false;
                for (let j = 0;j < donnees_fichier_controle_professeur.length && !trouve ; j++)
                {
                  if (donnees_fichier_controle_professeur[j][0] == nom_prenom_eleve)
                  {
                    trouve = true;
                    donnees_fichier_controle_professeur[j][3] = commentaire[3];
                  }
                }
                if (!trouve)
                {
                  msg_log = "Erreur eleve non trouvé : " + nom_prenom_eleve + " dans le fichier de contrôle des notes du professeur ";
                  msg_log += le_professeur_referent + " pour la kholle : " + feuille_releve_notes.getName();
                  ecrit_log(LOG_ERROR,msg_log);
                }
              }
              if (notes_remplies.length == releve_notes.length && commentaires_remplis.length == releve_notes.length) 
              {
                const etat_validation = feuille_releve_notes.getRange('E6').getValue();
                if (etat_validation)
                  etat_validation_kholleurs.push({"kholleur":kholleur,"etat_validation":true}); 
                else
                  etat_validation_kholleurs.push({"kholleur":kholleur,"etat_validation":false}); 
              }
              else
                etat_validation_kholleurs.push({"kholleur":kholleur,"etat_validation":false}); 
            }
          }
          fichier_controle_professeur.getDataRange().setValues(donnees_fichier_controle_professeur);
          let au_moins_un_kholleur_non_valide = false;
          const feuille = fichier_controle_professeur.getActiveSheet();
          for (const kholleur_etat_validation of etat_validation_kholleurs)
          {
            let j = 3;
            while (feuille.getRange(j,1).getValue() != "")
            {
              if (feuille.getRange(j,2).getValue() == kholleur_etat_validation["kholleur"])
              {
                if (kholleur_etat_validation["etat_validation"])
                  feuille.getRange(j,3,1,2).setBackground(null);
                else
                {
                  feuille.getRange(j,3,1,2).setBackground(null);
                  au_moins_un_kholleur_non_valide = true;
                }
              }
              j++;
            }
          }
          if (!au_moins_un_kholleur_non_valide)
          {
            if (feuille.getRange("E6").getValue())
            {
              const etat_planning = feuille_planning.getRange(i+1,6).getValue();
              if (etat_planning != "Oui")
              {
                feuille_planning.getRange(i+1,6).setValue("Oui");
                ecrit_spooler_resultats_eleves(classe,discipline,date_kholle,id_nom_fichier_controle_professeur);
              }
            }
            else
              feuille_planning.getRange(i+1,6).setValue("Non");
          }
        }
      }
    }
  }
}

function ecrit_spooler_resultats_eleves(classe,discipline,date_kholle,id_nom_fichier_controle_professeur)
{
  const sh_spooler = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_SPOOLER_ENVOIS_RESULTATS_ELEVES);
  let donnees = sh_spooler.getDataRange().getValues().filter(row => row[3] == id_nom_fichier_controle_professeur);
  if (donnees.length == 0)
    sh_spooler.appendRow([classe,discipline,date_kholle,id_nom_fichier_controle_professeur]);
  else
  {
    msg_log = "Entrée déjà ecrite dans Spooler : " + classe + " " + discipline + " " + date_kholle;
    ecrit_log(LOG_WARNING,"ecrit_spooler_resultats_eleves",msg_log); 
  }

}

function maj_reponses_confirmation_kholleur()
{
  let reponses_confirmation_kholleur = sheetToAssociativeArray(FEUILLE_CONFIRMATION_DISPO_KHOLLEURS);
  const sh_confirmations_dispo_kholleurs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CONFIRMATION_DISPO_KHOLLEURS);
  for (let i = 0; i <  reponses_confirmation_kholleur.length ; i++)
  {
    if (reponses_confirmation_kholleur[i]["Réponse"] == "")
    {
      const confirmation = reponses_confirmation_kholleur[i];
      const date_kholle = formatDate(confirmation["Date khôlle proposée"]);
      const annee = confirmation["Année"];
      const discipline = confirmation["Discipline"];
      const le_kholleur =confirmation["Civilité khôlleur"] + " " + confirmation["Nom khôlleur"] + " " + confirmation["Prénom khôlleur"];
      const nom_onglet_planning = construit_nom_onglet_planning(annee,date_kholle,discipline);
      const sh_nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
      const nb_examinateurs = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
      const donnees_kholleurs = lit_donnees_kholleurs_feuille_planning(sh_nom_onglet_planning,nb_examinateurs);
      let ind_examinateur = 1;
      let no_ligne_kholleur = -1;
      while (ind_examinateur <= nb_examinateurs && no_ligne_kholleur == -1)
      {
        const row = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES+ind_examinateur).getValue();
        const kholleur = donnees_kholleurs[ind_examinateur-1][0];
        if (kholleur == le_kholleur)
          no_ligne_kholleur = row;
        else
          ind_examinateur++;
      }
      if (no_ligne_kholleur == -1)
      {
        msg_log = "Erreur kholleur non trouvé " + le_kholleur + " dans l'onglet " + nom_onglet_planning;
        ecrit_log(LOG_ERROR,"maj_reponses_confirmation_kholleur",msg_log);
      }
      else
      {
        try
        {
          const form = FormApp.openById(confirmation["id_formulaire"]);
          const formResponses = form.getResponses();
          for (const reponse of formResponses) 
          {
            const itemResponses = reponse.getItemResponses();
            for (const item of itemResponses)
            {
              //const question = item.getItem().getTitle();
              let answer = item.getResponse();
              answer = answer.substring(0,3);
              sh_confirmations_dispo_kholleurs.getRange(i+2,8).setValue(answer);
              switch (answer)
              {
                case "Oui":
                  sh_nom_onglet_planning.getRange(no_ligne_kholleur,2).setBackground(COULEUR_CONFIRMATION_DISPONIBILITE_KHOLLEUR_OUI);
                  break;
                case "Non":
                  sh_nom_onglet_planning.getRange(no_ligne_kholleur,2).setBackground(COULEUR_CONFIRMATION_DISPONIBILITE_KHOLLEUR_NON);
                  break;
                default:
                  msg_log = "Erreur réponse au formulaire de confirmation de disponibilité inconnue : " + answer;
                  ecrit_log(LOG_ERROR,"maj_reponses_confirmation_kholleur",msg_log);
              }
            }
          }
        }
        catch (err)
        {
          msg_log = "Erreur formulaire confirmation dispo kholleur  id_form : " + confirmation["id_formulaire"] + err;
          ecrit_log(LOG_ERROR,"maj_reponses_confirmation_kholleur",msg_log);
        }
      }
    }
  }
}

function maj_parametres_kholle()
{
  maj_disponibilites_kholleurs();
  maj_heures_debut_et_fin_kholle();
}

function maj_disponibilites_kholleurs()
{
  const sh_choix_kholleurs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CHOIX_KHOLLEURS);
  const donnees_params_kholle = sh_choix_kholleurs.getDataRange().getValues();
  const sh_temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CALCUL_TEMPORAIRES);
  const annee = sh_temp.getRange("B7").getValue();
  const date_kholle = sh_temp.getRange("B5").getValue();
  const discipline = sh_temp.getRange("B6").getValue();
  const kholleurs = charge_kholleurs_discipline(annee,date_kholle,discipline,false);
  const nb_examinateurs = donnees_params_kholle[2][1];
  for (let i = 0; i < nb_examinateurs ; i++)
  {
    const le_kholleur = donnees_params_kholle[8+i][0];
    for (kholleur of kholleurs)
    {
      const civilite_nom_prenom = kholleur["civilite"] + " " + kholleur["nom"] + " " + kholleur["prenom"];
      if (civilite_nom_prenom == le_kholleur)
      {
        donnees_params_kholle[8+i][1] = kholleur["disponibilités"];
        const heure_split = kholleur["disponibilités"].split(" - ");
        const heure_debut = heure_split[0];
        const hh = heure_debut.substring(0,2);
        const mm = heure_debut.substring(3,5);
        donnees_params_kholle[8+i][3] = hh+":"+mm;
        break;
      }
    }
  }
  sh_choix_kholleurs.getDataRange().setValues(donnees_params_kholle);
}

function maj_heures_debut_et_fin_kholle()
{
  const sh_choix_kholleurs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CHOIX_KHOLLEURS);
  const donnees_params_kholle = sh_choix_kholleurs.getDataRange().getValues();
  const nb_examinateurs = donnees_params_kholle[2][1];
  const pause = donnees_params_kholle[5][1];
  const duree_pause_apres_chaque_passage = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("duree_pause_apres_chaque_passage").getValue();
  const duree_pause_passages = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("duree_pause_passages").getValue();
  const  nb_passages_entre_deux_pauses
    = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nb_passages_entre_deux_pauses").getValue();
  if (nb_examinateurs=="")
    return;
  const duree_kholle = donnees_params_kholle[3][1];
  for (let i = 0; i < nb_examinateurs;i++)
  {
    const nb_etudiants_kholleur = donnees_params_kholle[8+i][2];
    const heure_premiere_kholle = donnees_params_kholle[8+i][3];
    if (isInteger(nb_etudiants_kholleur) && isValidHourMinute(heure_premiere_kholle))
    {
      const heure_fin_derniere_kholle = calcule_heure_fin_kholle(heure_premiere_kholle,duree_kholle,pause,duree_pause_apres_chaque_passage,duree_pause_passages,nb_passages_entre_deux_pauses,nb_etudiants_kholleur)
      sh_choix_kholleurs.getRange(9+i,5).setValue(heure_fin_derniere_kholle);
    }
  }
}

function envoie_mails_kholleurs_feuille_notation()
{
  const sh_mails_saisis = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_MAILS_SAISIE_NOTES_KHOLLEURS);
  let donnees = sheetToAssociativeArray(FEUILLE_MAILS_SAISIE_NOTES_KHOLLEURS);
  const date_courante = new Date();
  for (let i = 0; i < donnees.length ; i++)
  {
    if (donnees[i]["Heure envoyée"] == "")
    {
      const civilite = donnees[i]["Civilité khôlleur"];
      const nom = donnees[i]["Nom khôlleur"];
      const prenom = donnees[i]["Prénom khôlleur"];
      const id_feuille_notation = donnees[i]["id feuille notation"];
      const date_envoi = donnees[i]["Date envoi"];
      const nom_fichier_releve_notes = donnees[i]["Nom fichier"];
      if (date_courante <= date_envoi)
      {
         envoie_mail_kholleur_saisie_notes(civilite,nom,prenom,id_feuille_notation,nom_fichier_releve_notes);
         sh_mails_saisis.getRange(i+1,10).setValue(date_courante);
      }
    }
  }
}

function envoie_resultats_tous_eleves_auto_validation_professeur()
{
   const sh_spooler = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_SPOOLER_ENVOIS_RESULTATS_ELEVES);
   let donnees = sh_spooler.getDataRange().getValues();
   const parametre_envoi_resultats_etudiants = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("parametre_envoi_resultats_etudiants").getValue();
   for (let ind = 1; ind < donnees.length; ind++)
   {
    let donnee = donnees[ind];
    if (donnee[4] == "")
    {
      const classe = donnee[0];
      const discipline = donnee[1];
      const date_kholle = formatDate(donnee[2]);
      const id_fichier_notation = donnee[3];
      const feuille_notes = SpreadsheetApp.openById(id_fichier_notation);
      if (!feuille_notes)
        return;
      const donnees_notes = feuille_notes.getDataRange().getValues().slice(2);  
      for (const donnee of donnees_notes)
      {
        const nom_prenom_eleve = donnee[0];
        const nom_prenom_eleve_split = nom_prenom_eleve.split(" ");
        const nom_eleve = nom_prenom_eleve_split[0];
        const prenom_eleve = nom_prenom_eleve_split[1];
        const kholleur = donnee[1];
        const l_eleve = retourne_eleve(nom_eleve,prenom_eleve);
        if (l_eleve==-1)
          return;
        const note = donnee[2];
        const commentaire = donnee[3];
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
        if (envoi_mail && note.toString().toUpperCase() != "ABS")
          envoie_mail_eleve_resultat_kholle(kholleur,l_eleve,discipline,classe,date_kholle,note,commentaire);
      }
      sh_spooler.getRange(ind+1,5).setValue(new Date());
      const nb_mails_envoyes = donnees_notes.length + "/" + donnees_notes.length;
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
              return;
            }
          }
        }
        i++;
      }
    }
  }
}
