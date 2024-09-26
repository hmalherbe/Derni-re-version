function initialisations_completes()
{
  init_histo_kholleurs_et_eleves();
  efface_logs();
  supprime_tous_les_plannings();
  masque_onglets_secondaires([FEUILLE_PLANNING_KHOLLES]);
  efface_infos_formulaires();
  const sh_temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CALCUL_TEMPORAIRES);
  sh_temp.clearContents();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_MAILS_SAISIE_NOTES_KHOLLEURS);
  clearSheetExceptFirstRow(sh);
  const sh_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_PLANNING_KHOLLES);
  sh_planning.getRange("D2:G100").clearContent();
}

function masque_onglets_secondaires(liste_feuilles_a_afficher)
{
  for (const onglet of SpreadsheetApp.getActiveSpreadsheet().getSheets())
  {
    const autre_onglet = onglet.getName();
    if (!liste_feuilles_a_afficher.includes(autre_onglet))
    { 
      onglet.hideSheet();
    }
  }
}

function init_histo_kholleurs_et_eleves()
{
  init_histo_kholleurs();
  init_histo_eleves();
}
                                                                                                      
function init_histo_kholleurs()
{
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_HISTO_KHOLLEURS);
  clearSheetExceptFirstRow(sheet);
  charge_kholleurs();
  charge_classes();
  
  let no_ligne = 2;
  for (let ind_classe = 0; ind_classe < les_classes.length ; ind_classe++)
  {
    let la_classe = les_classes[ind_classe]["année"];
    for (let ind_kholleur = 0; ind_kholleur < les_kholleurs.length ; ind_kholleur++)
    {
      sheet.getRange(no_ligne,1).setValue(la_classe);
      sheet.getRange(no_ligne,2).setValue(les_kholleurs[ind_kholleur]["nom"]);
      sheet.getRange(no_ligne,3).setValue(les_kholleurs[ind_kholleur]["prenom"]);
      sheet.getRange(no_ligne,4).setValue(les_kholleurs[ind_kholleur]["civilite"]);
      sheet.getRange(no_ligne,5).setValue(les_kholleurs[ind_kholleur]["matiere"]);
      sheet.getRange(no_ligne,6).setValue(0);
      no_ligne++;
    }
  }
}

function init_histo_eleves()
{
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_HISTO_ELEVES);
  clearSheetExceptFirstRow(sheet);
  /*
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_HISTO_ELEVES);
  sheet.clearContents();
  sheet.getRange(1,1).setValue("Nom de l'élève");
  sheet.getRange(1,2).setValue("Prénom de l'élève");
  sheet.getRange(1,3).setValue("Classe");
  sheet.getRange(1,4).setValue("Civilité du khôlleur");
  sheet.getRange(1,5).setValue("Nom du khôlleur");
  sheet.getRange(1,6).setValue("Prénom du khôlleur");
  sheet.getRange(1,7).setValue("Date de la khôlle");
  sheet.getRange(1,8).setValue("Heure de la khôlle");
  sheet.getRange(1,9).setValue("Position de la khôlle");
  sheet.getRange(1,10).setValue("Discipline de la khôlle");
  sheet.getRange(1,11).setValue("Nom prénom élève");
  sheet.getRange(1,12).setValue("Civilité nom prénom khôlleur");
  sheet.getRange(1,13).setValue("Civilité nom prénom professeur référent");
  sheet.getRange(1,14).setValue("Horodatage");
  */
}

function calcule_historique_eleves_old(classe) 
{
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_HISTO_ELEVES);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  sheet.getRange(2,1,lastRow,lastColumn).sort([{column: 3, ascending: true}, {column: 1, ascending: true}, {column: 2, ascending: true}, {column: 7, ascending: true}]);
  let histo_eleves = [];
  let no_ligne = 2;
  let prenom_eleve_prec = "";
  let nom_eleve_prec = "";
  let histo_un_eleve = -1;
  let max_score_heure_passage = -1;
  while (sheet.getRange(no_ligne,1).getValue() != "")
  {
    //console.log(no_ligne)
    let la_classe = sheet.getRange(no_ligne,3).getValue();
    if (la_classe == classe)
    {
      let nom_eleve = sheet.getRange(no_ligne,1).getValue();
      let prenom_eleve = sheet.getRange(no_ligne,2).getValue();
      if (nom_eleve != nom_eleve_prec && prenom_eleve != prenom_eleve_prec)
      {
        if (nom_eleve_prec != "")
        {
          let score_heure_passage =  histo_un_eleve["score_heure_passage"] / histo_un_eleve["nb_passages"];
          histo_un_eleve["score_heure_passage"] = score_heure_passage;
          if (score_heure_passage > max_score_heure_passage)
            max_score_heure_passage = score_heure_passage;
          histo_eleves.push(histo_un_eleve);
        }
        nom_eleve_prec = nom_eleve;
        prenom_eleve_prec = prenom_eleve;
        histo_un_eleve={"nom_eleve":nom_eleve,
                        "prenom_eleve":prenom_eleve,
                        "score_heure_passage":0,
                        "nb_passages":0
                        };
      }
      histo_un_eleve["score_heure_passage"]+=sheet.getRange(no_ligne,9).getValue();
      histo_un_eleve["nb_passages"]++;
    }
    no_ligne++;
  }
  if (histo_un_eleve !=-1)
  {
    let score_heure_passage =  histo_un_eleve["score_heure_passage"] / histo_un_eleve["nb_passages"];
    histo_un_eleve["score_heure_passage"] = score_heure_passage;
    if (score_heure_passage > max_score_heure_passage)
      max_score_heure_passage = score_heure_passage;
    histo_eleves.push(histo_un_eleve);
  }
  return {"histo_eleves":histo_eleves,"max_score_heure_passage":max_score_heure_passage};
}

function test_calcule_historique_eleves()
{
  const histo_eleves = calcule_historique_eleves("L1");
  //console.log("hm");
}

function calcule_historique_eleves(classe) 
{
  //creee_tcd(FEUILLE_HISTO_ELEVES,"A:N",FEUILLE_TCD_HISTO_ELEVES);
  let donnees_histo_eleves = sheetToAssociativeArray(FEUILLE_TCD_HISTO_ELEVES);
  if (donnees_histo_eleves.length==1)
    return {"histo_eleves":[],"max_score_heure_passage":-1};
  try
  {
    donnees_histo_eleves.filter(row => row["Classe -- Nom prenom élève"].startsWith(classe));
  }
  catch (err)
  {
    return {"histo_eleves":[],"max_score_heure_passage":-1};
  }
  
  let histo_eleves = [];
  let max_score_heure_passage = -1;
  for (const donnee of donnees_histo_eleves)
  {
      const classe_nom_prenom_eleve = donnee["Classe -- Nom prenom élève"];
      const classe_nom_prenom_eleve_split = classe_nom_prenom_eleve.split(" -- ");
      const nom_prenom_eleve = classe_nom_prenom_eleve_split[1];
      const nom_prenom_eleve_split = nom_prenom_eleve.split(" ");
      const nom_eleve = nom_prenom_eleve_split[0];
      const prenom_eleve = nom_prenom_eleve_split[1];
      const score_heure_passage = donnee["Moyenne position passage"];
      const nb_passages = donnee["Nb passages"];
      const histo_un_eleve={"nom_eleve":nom_eleve,
                          "prenom_eleve":prenom_eleve,
                          "score_heure_passage":score_heure_passage,
                          "nb_passages":nb_passages
                          };
      histo_eleves.push(histo_un_eleve);
      if (score_heure_passage > max_score_heure_passage)
        max_score_heure_passage = score_heure_passage;
    }
  return {"histo_eleves":histo_eleves,"max_score_heure_passage":max_score_heure_passage};
}

function lireDonneesTCDAvecQuery() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_TCD_HISTO_ELEVES);
  var formule = "=QUERY(A1:D10, \"select Col1, sum(Col2) group by Col1\")"; // Exemple de formule QUERY
  var resultat = sheet.getRange(sheet.getLastRow() + 1, 1).setFormula(formule).getValue();
  Logger.log(resultat);
}

function calcule_historique_kholleurs_old() 
{
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_HISTO_ELEVES);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  sheet.getRange(2,1,lastRow,lastColumn).sort([ {column: 5, ascending: true}, {column: 6, ascending: true}, {column: 7, ascending: true}]);
  let histo_kholleurs = [];
  let no_ligne = 2;
  let prenom_kholleur_prec = "";
  let nom_kholleur_prec = "";
  let histo_un_kholleur = -1;
  let date_kholle_prec = "";
  while (sheet.getRange(no_ligne,1).getValue() != "")
  {
    //console.log(no_ligne);
    let nom_kholleur = sheet.getRange(no_ligne,5).getValue();
    let prenom_kholleur = sheet.getRange(no_ligne,6).getValue();
    if (nom_kholleur != nom_kholleur_prec && prenom_kholleur != prenom_kholleur_prec)
    {
      if (nom_kholleur_prec != "")
      {
        for (let i = 0;i < histo_un_kholleur["histo_eleves"].length ; i++)
        {
          if (histo_un_kholleur["histo_eleves"][i]["nb_passages"] < histo_un_kholleur["min_passages_eleves"] || histo_un_kholleur["min_passages_eleves"] == -1)
            histo_un_kholleur["min_passages_eleves"] = histo_un_kholleur["histo_eleves"][i]["nb_passages"];
          if (histo_un_kholleur["histo_eleves"][i]["nb_passages"] < histo_un_kholleur["max_passages_eleves"] || histo_un_kholleur["max_passages_eleves"] == -1)
            histo_un_kholleur["max_passages_eleves"] = histo_un_kholleur["histo_eleves"][i]["nb_passages"];
        }
        histo_kholleurs.push(histo_un_kholleur);
      }
      nom_kholleur_prec = nom_kholleur;
      prenom_kholleur_prec = prenom_kholleur;
      histo_un_kholleur={"nom_kholleur":nom_kholleur,
                      "prenom_kholleur":prenom_kholleur,
                      "nb_passages":0,
                      "histo_eleves": [],
                      "min_passages_eleves":-1,
                      "max_passages_eleves":-1
                      };
      date_kholle_prec = "";
    }
    let date_kholle = sheet.getRange(no_ligne,7).getValue();
    let date_kholle_texte = date_kholle.getFullYear() + "-" + date_kholle.getMonth() + "-" + date_kholle.getDate();
    if (date_kholle_prec != date_kholle_texte)
    {
      histo_un_kholleur["nb_passages"]++;
      date_kholle_prec = date_kholle_texte;
    }
    let nom_eleve = sheet.getRange(no_ligne,1).getValue();
    let prenom_eleve = sheet.getRange(no_ligne,6).getValue();
    let deja = false;
    for (let ind_eleve=0;ind_eleve < histo_un_kholleur["histo_eleves"].length && !deja; ind_eleve++)
    {
      let un_eleve = histo_un_kholleur["histo_eleves"][ind_eleve];
      if (un_eleve["nom"] == nom_eleve && un_eleve["prenom"] == prenom_eleve)
      {
        histo_un_kholleur["histo_eleves"][ind_eleve]["nb_passages"]++;
      }
    }
    if (!deja)
    {
      const l_eleve = {"nom":nom_eleve,"prenom":prenom_eleve,"nb_passages":1};
      histo_un_kholleur["histo_eleves"].push(l_eleve);
    }
    no_ligne++;
  }
  if (histo_un_kholleur != -1)
  {
    for (let i = 0;i < histo_un_kholleur["histo_eleves"].length ; i++)
    {
      if (histo_un_kholleur["histo_eleves"][i]["nb_passages"] < histo_un_kholleur["min_passages_eleves"] || histo_un_kholleur["min_passages_eleves"] == -1)
        histo_un_kholleur["min_passages_eleves"] = histo_un_kholleur["histo_eleves"][i]["nb_passages"];
      if (histo_un_kholleur["histo_eleves"][i]["nb_passages"] < histo_un_kholleur["max_passages_eleves"] || histo_un_kholleur["max_passages_eleves"] == -1)
        histo_un_kholleur["max_passages_eleves"] = histo_un_kholleur["histo_eleves"][i]["nb_passages"];
    }
    histo_kholleurs.push(histo_un_kholleur);
  }
  return histo_kholleurs;
}

function calcule_historique_kholleurs() 
{
  let donnees_histo_kholleurs_eleves = sheetToAssociativeArray(FEUILLE_TCD_HISTO_KHOLLEURS_ELEVES);
  if (donnees_histo_kholleurs_eleves.length == 1)
    return [];
  let histo_kholleurs = [];
  let histo_un_kholleur = -1;
  let kholleur_prec = "";
  let max_passages = -1;
  let nom_kholleur;
  let prenom_kholleur;
  for (const donnee of donnees_histo_kholleurs_eleves)
  {
    const kholleur_eleve = donnee["Civilité nom prénom khôlleur -- Nom prénom élève"];
    const kholleur_eleve_split  = kholleur_eleve.split(" -- ");
    const kholleur = kholleur_eleve_split[0];
    const nb_passages = donnee["Nb Passages"];
    if (kholleur != kholleur_prec)
    {
      if (kholleur_prec != "")
      {
        histo_un_kholleur["max_passages_eleves"] = max_passages;
        histo_un_kholleur["nb_passages"] = retourne_nb_passages_histo_kholleur(nom_kholleur,prenom_kholleur);
        histo_kholleurs.push(histo_un_kholleur);
      }
      kholleur_prec = kholleur;
      const kholleur_split = kholleur.split(" ");
      nom_kholleur = kholleur_split[1];
      prenom_kholleur = kholleur_split[2];
      histo_un_kholleur={"nom_kholleur":nom_kholleur,
                      "prenom_kholleur":prenom_kholleur,
                      "histo_eleves": []
                      };
      max_passages = -1;
      somme_passages = 0;
    }
    const eleve = kholleur_eleve_split[1];
    const eleve_split = eleve.split(" ");
    const nom_eleve = eleve_split[0];
    const prenom_eleve = eleve_split[1];
    const l_eleve = {"nom":nom_eleve,"prenom":prenom_eleve,"nb_passages":nb_passages};
    histo_un_kholleur["histo_eleves"].push(l_eleve);
    somme_passages += nb_passages;
    if (nb_passages > max_passages)
      max_passages = nb_passages;
  }
  if (kholleur_prec != "")
  {
    histo_un_kholleur["max_passages_eleves"] = max_passages;
    histo_kholleurs.push(histo_un_kholleur);
  }
  return histo_kholleurs;
}

function retourne_nb_passages_histo_kholleur(nom_kholleur,prenom_kholleur)
{ 
  let donnees_histo_kholleurs = sheetToAssociativeArray(FEUILLE_HISTO_KHOLLEURS);
  donnees_histo_kholleurs = donnees_histo_kholleurs.filter(row => row["Nom Kholleur"] == nom_kholleur && row["Prénom Kholleur"] == prenom_kholleur);
  if (donnees_histo_kholleurs.length == 0)
  {
    msg_log = "Erreur kholleur non trouvé dans histo : " + nom_kholleur + " " + prenom_kholleur + " ("+ FEUILLE_HISTO_KHOLLEURS+")";
    ecrit_log(LOG_ERROR,"retourne_nb_passages_histo_kholleur",msg_log);
    return -1;
  }
  else
  {
    let nb_passages = 0;
    for (const donnee of donnees_histo_kholleurs)
      nb_passages += parseInt(donnee["Nb passages"]);
    return nb_passages;
  }
}


function retourne_histo_kholleur(prenom_kholleur,nom_kholleur,histo_kholleurs)
{
  for (let i = 0;i < histo_kholleurs.length;i++)
  {
    if (prenom_kholleur == histo_kholleurs[i]["prenom_kholleur"] && nom_kholleur == histo_kholleurs[i]["nom_kholleur"])
      return histo_kholleurs[i];
  }
  return -1;
}

function enregistre_seance_histo_kholleurs()
{
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  const sh_nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  const infos_planning = retourne_infos_from_nom_onglet_planning(nom_onglet_planning);
  const annee = infos_planning["annee"];
  const discipline = infos_planning["discipline"];
  const nb_examinateurs = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
  for (let ind_examinateur = 1; ind_examinateur <= nb_examinateurs; ind_examinateur++)
  {
    const row = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES+ind_examinateur).getValue();
    const nom_kholleur = sh_nom_onglet_planning.getRange(row,2).getValue();
    if (nom_kholleur != "")
    {
      maj_histo_kholleurs(annee,discipline,nom_kholleur);
      msg_sideBar += "<BR><BR> Traitement de l'examinateur : " + nom_kholleur;
      SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue(msg_sideBar);
      showSideBarInfos("Enregistrement du planning de khôlle");
    }
      
    else
      return null;
  }
  return nb_examinateurs;
}

function maj_histo_kholleurs(annee,discipline,nom_kholleur)
{
  let sh_histo_kholleurs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_HISTO_KHOLLEURS);
  let nb_kholleurs = 0;
  let trouve = false;
  while (sh_histo_kholleurs.getRange(nb_kholleurs+2,1).getValue() != "" && !trouve)
  {
    let la_discipline = sh_histo_kholleurs.getRange(nb_kholleurs+2,5).getValue();
    let l_annee = sh_histo_kholleurs.getRange(nb_kholleurs+2,1).getValue();
    let la_civilite = sh_histo_kholleurs.getRange(nb_kholleurs+2,4).getValue();
    let le_prenom = sh_histo_kholleurs.getRange(nb_kholleurs+2,3).getValue();
    let le_nom = sh_histo_kholleurs.getRange(nb_kholleurs+2,2).getValue();
    let le_kholleur = formate_nom_kholleur(la_civilite,le_prenom,le_nom);
    if (discipline == la_discipline && annee == l_annee && nom_kholleur == le_kholleur)
    {
      let nb_passages = sh_histo_kholleurs.getRange(nb_kholleurs+2,6).getValue()+1;
      sh_histo_kholleurs.getRange(nb_kholleurs+2,6).setValue(nb_passages);
      trouve = true;
    }
    nb_kholleurs++;
  }
  if (!trouve)
  {
    sh_histo_kholleurs.getRange(nb_kholleurs+2,1).setValue(annee);  
    let split_nom_kholleur = nom_kholleur.split(" ");
    let civilite = split_nom_kholleur[0];
    let prenom = split_nom_kholleur[1];
    let nom = split_nom_kholleur[2];
    sh_histo_kholleurs.getRange(nb_kholleurs+2,2).setValue(nom);
    sh_histo_kholleurs.getRange(nb_kholleurs+2,3).setValue(prenom);
    sh_histo_kholleurs.getRange(nb_kholleurs+2,4).setValue(civilite);
    sh_histo_kholleurs.getRange(nb_kholleurs+2,5).setValue(discipline);
    sh_histo_kholleurs.getRange(nb_kholleurs+2,6).setValue(1);
  }
}

function determine_no_kholleur(no_eleve,repart_nb_eleves_par_kholleur)
{
  let cp_repart = [...repart_nb_eleves_par_kholleur];
  for (let i = 0; i < cp_repart.length-1; i++)
  {
    cp_repart[i+1] += cp_repart[i];
  }
  for (let i = 0; i < cp_repart.length; i++)
  {
    if (no_eleve < cp_repart[i])
      return i+1;
  }
  msg_log = "Erreur dans determine_no_kholleur";
  ecrit_log(LOG_ERROR,"determine_no_kholleur",msg_log);
  return -1;
}

function enregistre_seance_histo_eleves()
{
  let sh_histo_eleves = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_HISTO_ELEVES);
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  let sh_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  let no_eleve = 0;
  let no_ligne = 3;
  let ligne = [];
  const infos_planning = retourne_infos_from_nom_onglet_planning(nom_onglet_planning)
  const classe = infos_planning["annee"];
  const date_kholle = infos_planning["date_kholle"];
  const discipline = infos_planning["discipline"];
  const nb_examinateurs = sh_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
  const nb_eleves = sh_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES).getValue();
  const donnees_kholleurs = lit_donnees_kholleurs_feuille_planning(sh_planning,nb_examinateurs);
  let repart_nb_eleves_par_kholleur = [];
  for (let i = 0; i < nb_examinateurs ; i++)
  {
    repart_nb_eleves_par_kholleur[i] = donnees_kholleurs[i][1];
  }
  const professeur_referent = sh_planning.getRange(1,8).getValue();
  let position_kholle=0;
  let no_kholleur_prec = -1;
  let liste_eleves_meme_kholleur = [];
  while (sh_planning.getRange(no_ligne,3).getValue() != "")
  {
    if (sh_planning.getRange(no_ligne,3).getValue().substr(0,5) != 'Pause')
    {
      let nom_prenom_eleve = sh_planning.getRange(no_ligne,3).getValue();
      if (nom_prenom_eleve != "")
      {
        let split_nom_prenom_eleve = nom_prenom_eleve.split(" ");
        let nom_eleve = split_nom_prenom_eleve[0];
        let prenom_eleve = split_nom_prenom_eleve[1];
        let no_kholleur =determine_no_kholleur(no_eleve,repart_nb_eleves_par_kholleur);
        const heure_preparation = sh_planning.getRange(no_ligne,4).getValue();
        const heure_passage = sh_planning.getRange(no_ligne,5).getValue();
        if (no_kholleur_prec == no_kholleur)
        {
          liste_eleves_meme_kholleur.push({"nom":nom_eleve,"prénom":prenom_eleve,"heure_preparation":heure_preparation,"heure_passage":heure_passage});
          position_kholle++;
        }
        else
        {
          if (no_kholleur_prec != -1)
          {
            const row = sh_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES+no_kholleur_prec).getValue();
            const derniere_plage = sh_planning.getRange(no_ligne-1,5).getValue();
            const derniere_plage_split = derniere_plage.split(" - ");
            const derniere_heure = derniere_plage_split[1];
            const salle = sh_planning.getRange(row,1).getValue();
             if (AVEC_CLASSROOM)
              cree_kholle_classroom(civilite_nom_prenom_kholleur_dict,classe,date_kholle,discipline,liste_eleves_meme_kholleur);
             else
              cree_feuille_notation(no_kholleur_prec,civilite_nom_prenom_kholleur_dict,classe,date_kholle,discipline,liste_eleves_meme_kholleur,salle,derniere_heure);
          }
          liste_eleves_meme_kholleur = [{"nom":nom_eleve,"prénom":prenom_eleve,"heure_preparation":heure_preparation,"heure_passage":heure_passage}];
          position_kholle = 1;
          no_kholleur_prec = no_kholleur;
        }
        const civilite_nom_prenom_kholleur = donnees_kholleurs[no_kholleur-1][0];
        let split_civilite_nom_prenom_kholleur = civilite_nom_prenom_kholleur.split(" ");
        let civilite_kholleur = split_civilite_nom_prenom_kholleur[0];
        let nom_kholleur = split_civilite_nom_prenom_kholleur[1];
        let prenom_kholleur = split_civilite_nom_prenom_kholleur[2];
        let heure_kholle = sh_planning.getRange(no_ligne,5).getValue();
        const classe_eleve = classe + " -- " + nom_prenom_eleve;
        const kholleur_eleve = civilite_nom_prenom_kholleur + " -- " + nom_prenom_eleve;
        var civilite_nom_prenom_kholleur_dict = {"civilité":civilite_kholleur,"nom":nom_kholleur,"prénom":prenom_kholleur};
        ligne = [nom_eleve,prenom_eleve,classe,civilite_kholleur,nom_kholleur,prenom_kholleur,date_kholle,heure_kholle,position_kholle,discipline,nom_prenom_eleve,civilite_nom_prenom_kholleur,professeur_referent,classe_eleve,kholleur_eleve,new Date()];
        sh_histo_eleves.appendRow(ligne);
      }
      no_eleve++;
    }
    no_ligne++;
  } 
  if (no_kholleur_prec != -1)
  {
     if (AVEC_CLASSROOM)
        cree_kholle_classroom(civilite_nom_prenom_kholleur_dict,classe,date_kholle,discipline,liste_eleves_meme_kholleur);
     else
     {
        const row = sh_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES+no_kholleur_prec).getValue();
        const salle = sh_planning.getRange(row,1).getValue();
        const derniere_plage = sh_planning.getRange(no_ligne-1,5).getValue();
        cree_feuille_notation(no_kholleur_prec,civilite_nom_prenom_kholleur_dict,classe,date_kholle,discipline,liste_eleves_meme_kholleur,salle,derniere_plage);
     }
  }
  return no_eleve;
}

function retourne_nb_passages_kholleurs(annee,nom,prenom)
{
 let sh_histo_kholleurs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_HISTO_KHOLLEURS);
 let i = 2; 
 while (sh_histo_kholleurs.getRange(i,1).getValue() != "")
 {
    if (annee == sh_histo_kholleurs.getRange(i,1).getValue() && 
        nom == sh_histo_kholleurs.getRange(i,2).getValue()   &&
        prenom == sh_histo_kholleurs.getRange(i,3).getValue())
        return sh_histo_kholleurs.getRange(i,6).getValue();
    i++;
 }
 console.log("Erreur retourne_nb_passages_kholleurs : kholleur non reconnu " + nom + " - " + prenom);
}