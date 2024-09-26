const JOURS_KOLLES = {"L1":[{"lundi":"18:00"},{"mardi":"18:00"},{"jeudi":"18:00"}],
                      "L2":[{"lundi":"18:10"},{"mercredi":"10:10"}]
        }
const PAUSES = ["5 min après chaque passage","10 min tous les 4 passages","Pas de pause"];	
const NB_EXAMINATEURS = [3,4];

function genere_dispos_aleatoires_old() 
{
  let classeur;
  if (AVEC_ANNUAIRE_PARTAGE)
    classeur = SpreadsheetApp.openById(ID_ANNUAIRE);
  else
    classeur = SpreadsheetApp.getActiveSpreadsheet();
  let sh_kholleurs = classeur.getSheetByName(FEUILLE_LISTE_KHOLLEURS);
  let i = 3;
  while (sh_kholleurs.getRange(i,1).getValue() != "")
  {
    for (let j = 0; j <= 4; j++)
    {
      if (Math.random() < 0.5)
        sh_kholleurs.getRange(i,6+j).setValue("TRUE");
      else
        sh_kholleurs.getRange(i,6+j).setValue("FALSE");
    }
    i++;
  }
}

const PLAGES_DISPO = ["18:00 - 22:00","18:00 - 21:00","18:00 - 20:00","19:00 - 22:00","19:00 - 21:00","19:00 - 20:00","20:00 - 22:00","20:00 - 21:00"];

function genere_dispos_aleatoires(taux_remplissage = 0.2) 
{
  let classeur;
  if (AVEC_ANNUAIRE_PARTAGE)
    classeur = SpreadsheetApp.openById(ID_ANNUAIRE);
  else
    classeur = SpreadsheetApp.getActiveSpreadsheet();
  let sh_kholleurs = classeur.getSheetByName(FEUILLE_LISTE_KHOLLEURS);
  let i = 3;
  while (sh_kholleurs.getRange(i,1).getValue() != "")
  {
    for (let j = 0; j <= 4; j++)
    {
      if (Math.random() < taux_remplissage)
        sh_kholleurs.getRange(i,6+j).setValue("");
      else
        sh_kholleurs.getRange(i,6+j).setValue(shuffle(PLAGES_DISPO)[0]);
    }
    i++;
  }
}


function genere_evaluation_aleatoire_kholleurs() 
{
  let classeur;
  if (AVEC_ANNUAIRE_PARTAGE)
    classeur = SpreadsheetApp.openById(ID_ANNUAIRE);
  else
    classeur = SpreadsheetApp.getActiveSpreadsheet();
  let sh_kholleurs = classeur.getSheetByName(FEUILLE_LISTE_KHOLLEURS);
  let i = 3;
  while (sh_kholleurs.getRange(i,1).getValue() != "")
  {
      sh_kholleurs.getRange(i,11).setValue(Math.floor(Math.random() * 5)+1);
    i++;
  }
}

function genere_plannings()
{
  charge_matieres();
  let les_matieres_non_lv = les_matieres.filter(row => row["LV"]=="Non");
  const les_matieres_lv = les_matieres.filter(row => row["LV"]=="Oui");
  const calendrier = charge_calendrier();
  //createOrReplaceSheet(FEUILLE_PLANNING_TEST);
  const sh_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_PLANNING_KHOLLES);
  for (const date of calendrier)
  {
    if (date["Semaine Active"] == "Oui")
    {
      les_matieres_non_lv = shuffle(les_matieres_non_lv);
      let no_matiere = 0;
      for (let no_jour=0;no_jour<6;no_jour++)
      {
        let jour_semaine = JOURS_SEMAINE[no_jour];
        for (const annee of ["L2"])
        {
          for (const jours_kholles of JOURS_KOLLES[annee])
          {
            if (jour_semaine in jours_kholles)
            {
              let datek = date["Date"];
              const date_kholle = new Date(datek.getFullYear(),datek.getMonth(),datek.getDate() + no_jour);
              const discipline = les_matieres_non_lv[no_matiere % les_matieres_non_lv.length]["nom"];
              no_matiere++;
              sh_planning.appendRow([annee,date_kholle,discipline]);
            }
         }
        }
      }
    }
  }
}

function charge_calendrier()
{
  return sheetToAssociativeArray(FEUILLE_CALENDRIER);

  
  console("hm");
}


function tests_tous_les_plannings()
{
  initialisations_completes();
  genere_dispos_aleatoires();
  const donnees_plannings = sheetToAssociativeArray(FEUILLE_PLANNING_KHOLLES);
  let heure_debut;
  const sh_choix =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CHOIX_KHOLLEURS);
  for (let i = 0; i < donnees_plannings.length ; i++)
  {
    if (donnees_plannings[i]["Planning enregistré"] != "Oui")
    {
      console.log((i+1)+"/"+donnees_plannings.length + " "+ formatDate(donnees_plannings[i]["Date"]) + " " + donnees_plannings[i]["Année"] + " " + donnees_plannings[i]["Discipline"]);
    choisit_le_planning(i);
    proposer_kholleurs_avec_dispo();
    ecrire_planning_avec_kholleurs_choisis();
    proposer_eleves_V2();
    enregistre_seance();
    }
  }
}
