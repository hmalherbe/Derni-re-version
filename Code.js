function onOpen() {
  var menu_initialisation = [ /*{name: 'Initialiser Classroom', functionName: 'init_classroom'},*/
                {name: 'Initialiser Histo khôlleurs et élèves', functionName: 'init_histo_kholleurs_et_eleves'},
                {name: 'Initialiser disponibilités khôlleurs', functionName: 'init_dispo_kholleurs'},
                {name: 'Effacer les logs', functionName: 'efface_logs'},
                {name: 'Effacer les infos des formulaires', functionName: 'efface_infos_formulaires'},
                {name: 'Initialisations complètes', functionName: 'initialisations_completes'}
                ];
  var menu_tests = [ 
                {name: 'Simuler tous les plannings', functionName: 'tests_tous_les_plannings'}
                ];
  var menu_disponibilites_kholleurs = [ {name: 'Annuaire Etudiants - Examinateurs - Professeurs référents', functionName: 'consulter_maj_annuaire'}];
  var menu_planning = [ {name: 'Afficher le planning des khôlles', functionName: 'affiche_planning_interrogations'},
                        {name: 'Choisir un planning de khôlle', functionName: 'choisit_planning_interrogations'},
                        {name: 'Proposer des examinateurs', functionName: 'proposer_kholleurs_avec_dispo'},
                        //{name: 'Proposer des examinateurs (sans disponibilité)', functionName: 'proposer_kholleurs_sans_dispo'},
                         {name: 'Ecrire le planning avec les examinateurs choisis', functionName: 'ecrire_planning_avec_kholleurs_choisis'},
                        {name: 'Proposer les étudiants', functionName: 'proposer_eleves_V2'},
                        {name: 'Demander la confirmation de disponibilités des examinateurs proposés', functionName: 'demande_confirmation_disponibilite_kholleur'},
                        {name: 'Valider la grille', functionName: 'enregistre_seance'},               
                ];

  var menu_mailings = [ {name: 'Renvoyer le mail avec le relevé des notes et commentaires pour les examinateurs', functionName: 'envoyer_mail_rappel_kholleurs_feuille_notation'},
                        {name: 'Renvoyer le mail  avec le contrôle des notes et commentaires pour le professeur référent', functionName: 'envoyer_mail_rappel_professeur_controle_notation'},
                        {name: 'Envoyer les notes / commentaires aux étudiants', functionName: 'envoyer_notes_commentaires_eleves'}           
                ];

  SpreadsheetApp.getActive().addMenu('Initialisations', menu_initialisation);
  SpreadsheetApp.getActive().addMenu('Annuaire', menu_disponibilites_kholleurs);
  SpreadsheetApp.getActive().addMenu('Générer le planning des khôlles', menu_planning);
   SpreadsheetApp.getActive().addMenu('Mailings', menu_mailings);
   SpreadsheetApp.getActive().addMenu('Tests', menu_tests);
}

function menuperso() {
  var menu_initialisation = [ 
                {name: 'Initialiser Histo khôlleurs et élèves', functionName: 'init_histo_kholleurs_et_eleves'},
                {name: 'Initialiser disponibilités khôlleurs', functionName: 'init_dispo_kholleurs'},
                {name: 'Effacer les logs', functionName: 'efface_logs'},
                {name: 'Effacer les infos des formulaires', functionName: 'efface_infos_formulaires'}
                ];

  SpreadsheetApp.getActive().addMenu('Initialisations', menu_initialisation);

}

var les_cours;
var les_matieres = [];
var les_classes = [];
var les_kholleurs = [];
var les_professeurs = [];
var les_eleves = [];
var les_salles_kholles = [];
var sh_param;
var msg_log = "";

function consulter_maj_annuaire()
{
  let classeur;
  if (AVEC_ANNUAIRE_PARTAGE)
  {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('Annuaire_profs_kholleurs_etudiants').setTitle("Annuaire 2024-2025");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
  }
  else
  {
    classeur = SpreadsheetApp.getActiveSpreadsheet();
    classeur.getSheetByName(FEUILLE_LISTE_KHOLLEURS).activate();
  }
    
}

function efface_logs()
{
  let sh_logs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_LOGS);
  clearSheetExceptFirstRow(sh_logs);
}

function efface_infos_formulaires()
{
   let sh_form = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CONFIRMATION_DISPO_KHOLLEURS);
   clearSheetExceptFirstRow(sh_form);
}

function charge_kholleurs()
{
  let classeur;
  if (AVEC_ANNUAIRE_PARTAGE)
    classeur = SpreadsheetApp.openById(ID_ANNUAIRE);
  else
    classeur = SpreadsheetApp.getActiveSpreadsheet();
  let sh_kholleurs = classeur.getSheetByName(FEUILLE_LISTE_KHOLLEURS);
  let nb_kholleurs = 0;
  les_kholleurs = [];
  while (sh_kholleurs.getRange(nb_kholleurs+3,1).getValue() != "")
  {
    let kholleur = {"nom":sh_kholleurs.getRange(nb_kholleurs+3,3).getValue(),
                    "prenom":sh_kholleurs.getRange(nb_kholleurs+3,4).getValue(),
                    "civilite":sh_kholleurs.getRange(nb_kholleurs+3,2).getValue(),
                    "mail":sh_kholleurs.getRange(nb_kholleurs+3,colonne_mail_kholleur).getValue(),
                    "matiere":sh_kholleurs.getRange(nb_kholleurs+3,1).getValue(),"identifiant_dossier_notes":sh_kholleurs.getRange(nb_kholleurs+3,1).getValue(),"evaluation":sh_kholleurs.getRange(nb_kholleurs+3,11).getValue(),"cours_assignes":[]}
    let dispos = {};
    for (let i = 1; i <= 5; i++)
    {
      dispos[JOURS_SEMAINE[i-1]] = sh_kholleurs.getRange(nb_kholleurs+3,5+i).getValue();
    }
    kholleur["disponibilites"] = dispos;
    les_kholleurs.push(kholleur);
    nb_kholleurs++;
  }
}

function charge_professeurs()
{
  let classeur;
  if (AVEC_ANNUAIRE_PARTAGE)
    classeur = SpreadsheetApp.openById(ID_ANNUAIRE);
  else
    classeur = SpreadsheetApp.getActiveSpreadsheet();
  let sh_professeurs = classeur.getSheetByName(FEUILLE_LISTE_PROFESSEURS);
  let nb_professeurs = 0;
  les_professeurs = [];
  while (sh_professeurs.getRange(nb_professeurs+2,1).getValue() != "")
  {
    let professeur = {  "année":sh_professeurs.getRange(nb_professeurs+2,1).getValue(),
                        "matière":sh_professeurs.getRange(nb_professeurs+2,2).getValue(),
                        "nom":sh_professeurs.getRange(nb_professeurs+2,3).getValue(),
                        "prénom":sh_professeurs.getRange(nb_professeurs+2,4).getValue(),"civilité":sh_professeurs.getRange(nb_professeurs+2,5).getValue(),
                        "mail":sh_professeurs.getRange(nb_professeurs+2,colonne_mail_prof).getValue(),
                        "semestre":sh_professeurs.getRange(nb_professeurs+2,7).getValue(),"complément":sh_professeurs.getRange(nb_professeurs+2,8).getValue(),"cours_assignes":[]}
    les_professeurs.push(professeur);
    nb_professeurs++;
  }
}



function charge_kholleurs_discipline(annee,date_kholle,discipline,avec_dispo_kholleurs)
{
  let classeur;
  if (AVEC_ANNUAIRE_PARTAGE)
    classeur = SpreadsheetApp.openById(ID_ANNUAIRE);
  else
    classeur = SpreadsheetApp.getActiveSpreadsheet();
  let sh_kholleurs = classeur.getSheetByName(FEUILLE_LISTE_KHOLLEURS);
  let i = 3;
  let les_kholleurs_discipline = [];
  while (sh_kholleurs.getRange(i,1).getValue() != "")
  {
    const nom = sh_kholleurs.getRange(i,3).getValue();
    const prenom = sh_kholleurs.getRange(i,4).getValue();
    const civilite = sh_kholleurs.getRange(i,2).getValue();
    const matiere = sh_kholleurs.getRange(i,1).getValue();
    const evaluation = sh_kholleurs.getRange(i,11).getValue();
    if (matiere == discipline)
    {
      let dispo = sh_kholleurs.getRange(i,5+date_kholle.getDay()).getValue();
      if (!avec_dispo_kholleurs || dispo) 
      {
        let nb_passages = retourne_nb_passages_kholleurs(annee,nom,prenom);
        let le_kholleur = {"civilite":civilite,"nom":nom,"prenom":prenom,"nb_passages":nb_passages,"evaluation":evaluation,"disponibilités":dispo};
        les_kholleurs_discipline.push(le_kholleur);
      }
    }
    i++;
  }
  return les_kholleurs_discipline;
}

function charge_professeurs_discipline(annee,discipline)
{
  let classeur;
  if (AVEC_ANNUAIRE_PARTAGE)
    classeur = SpreadsheetApp.openById(ID_ANNUAIRE);
  else
    classeur = SpreadsheetApp.getActiveSpreadsheet();
  let sh_professeurs = classeur.getSheetByName(FEUILLE_LISTE_PROFESSEURS);
  const donnees_professeurs = sh_professeurs.getDataRange().getValues();
  let donnees_professeurs_filtrees = donnees_professeurs.filter(function(row) 
    {
      //console.log(row[0],row[0] == annee,row[1],row[1] == discipline);
      return row[0] == annee && row[1] == discipline;
      //return true;
    });
  return donnees_professeurs_filtrees;
}

function charge_salles_kholles()
{
  const range_salles_kholles = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(PLAGE_SALLES_KHOLLES);
  les_salles_kholles = [];
  if (range_salles_kholles) {
    const values = range_salles_kholles.getValues();
    for (let i = 0; i < values.length; i++) {
        les_salles_kholles.push(values[i][0]);
      }
    }
 else {
    // Handle the case where the named range doesn't exist
    console.error("Named range 'Salles_kholles' not found.");
  }
}

function charge_eleves()
{
  let classeur;
  if (AVEC_ANNUAIRE_PARTAGE)
    classeur = SpreadsheetApp.openById(ID_ANNUAIRE);
  else
    classeur = SpreadsheetApp.getActiveSpreadsheet();
  let sh_eleves = classeur.getSheetByName(FEUILLE_LISTE_ELEVES);
  let nb_eleves = 0;
  les_eleves = [];
  while (sh_eleves.getRange(nb_eleves+2,1).getValue() != "")
  {
    let eleve = {"nom":sh_eleves.getRange(nb_eleves+2,1).getValue(),"prenom":sh_eleves.getRange(nb_eleves+2,2).getValue(),"mail":sh_eleves.getRange(nb_eleves+2,colonne_mail_eleve).getValue(),"classe":sh_eleves.getRange(nb_eleves+2,4).getValue(),"lv":sh_eleves.getRange(nb_eleves+2,5).getValue(),"cours_assignes":[]}
    les_eleves.push(eleve);
    nb_eleves++;
  }
}

function charge_eleves_annee()
{
  charge_eleves_annee_discipline("L2","Droit");
}

function charge_eleves_annee_discipline(annee,discipline,deja_eleves=[])
{
  let classeur;
  if (AVEC_ANNUAIRE_PARTAGE)
    classeur = SpreadsheetApp.openById(ID_ANNUAIRE);
  else
    classeur = SpreadsheetApp.getActiveSpreadsheet();
  let sh_eleves = classeur.getSheetByName(FEUILLE_LISTE_ELEVES);
  let nb_eleves = 0;
  let eleves = [];
  while (sh_eleves.getRange(nb_eleves+2,1).getValue() != "")
  {
    let la_classe = sh_eleves.getRange(nb_eleves+2,4).getValue();
    let LV = sh_eleves.getRange(nb_eleves+2,5).getValue();
    if (annee == la_classe && (!est_LV(discipline) || LV == discipline))
    {
      let prenom = sh_eleves.getRange(nb_eleves+2,2).getValue();
      let nom = sh_eleves.getRange(nb_eleves+2,1).getValue();
      let deja = false;
      for (deja_eleve of deja_eleves)
      {
        if (deja_eleve["prénom"] == prenom && deja_eleve["nom"] == nom)
        {
          deja = true;
          break;
        }
      }
      if (!deja)
      {
        let eleve = nom + " " + prenom;
        eleves.push(eleve);
      }
    }
    nb_eleves++;
  }
  //return shuffle(eleves);
  return eleves;
}

function charge_matieres()
{
  sh_param = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_PARAMETRES);
  let nb_matieres = 0;
  les_matieres = [];
  //console.log("hm " + sh_param.getRange(3,1).getValue());
  while (sh_param.getRange(nb_matieres+2,1).getValue() != "")
  {
    let matiere = { "nom":sh_param.getRange(nb_matieres+2,1).getValue(),
                    "LV":sh_param.getRange(nb_matieres+2,2).getValue(),
                    "duree_preparation":sh_param.getRange(nb_matieres+2,3).getValue(),
                    "annees":sh_param.getRange(nb_matieres+2,4).getValue()
                    };
    les_matieres.push(matiere);
    nb_matieres++;
    //break;
  }
}

function charge_classes()
{
  let sh_param = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_PARAMETRES);
  let nb_classes  = 0;
  les_classes = [];
  while (sh_param.getRange(nb_classes+2,6).getValue() != "")
  {
    let la_classe = {"nom":sh_param.getRange(nb_classes+2,6).getValue(),"année":sh_param.getRange(nb_classes+2,7
    ).getValue()};
    les_classes.push(la_classe);
    nb_classes++;
  }
}

function est_LV(matiere)
{
  if (les_matieres.length==0)
    charge_matieres();
  for (let i = 0;i < les_matieres.length; i++)
  {
    if (les_matieres[i].nom == matiere)
    {
      return les_matieres[i]["LV"]=="Oui";
    }
  }
  console.log("Pb matiere inconnue " + matiere);
}

function supprime_cours()
{
  const les_cours = Classroom.Courses.list();  
  let les_id_cours = [];
  let cours = les_cours.getCourses();
  let nb_cours_supprimes;
  if (!cours)
  {
    nb_cours_supprimes = 0;
    return;
  }
  nb_cours_supprimes = cours.length;
  for (let i = 0; i < cours.length; i++)
  {
    les_id_cours.push(cours[i].getId())
  }
  for (let i = 0; i < les_id_cours.length; i++)
  {
    Classroom.Courses.remove(les_id_cours[i]);
  }
  return nb_cours_supprimes;
}

function lit_preparation(discipline)
{
  for (let i = 0;i < les_matieres.length;i++)
  {
    if (les_matieres[i]["nom"] == discipline)
      return les_matieres[i]["duree_preparation"]
  }
  console.log("Erreur : matière non reconnue " + discipline);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('propose_planning');
}

function init_dispo_kholleurs()
{
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(feuille_dispo_kholleurs);
  sheet.clearContents();
  charge_kholleurs();
  let ligne = ["Nom khôlleur","Prénom khôlleur","Discipline","lundi","mardi","mercredi","jeudi","vendredi"];
  sheet.appendRow(ligne);
  for (let i = 0; i < les_kholleurs.length;i++)
  {
    let nom_kholleur = les_kholleurs[i]["nom"];
    let prenom_kholleur = les_kholleurs[i]["prenom"];
    let matiere = les_kholleurs[i]["matiere"];
    sheet.getRange(i+2,1).setValue(nom_kholleur);
    sheet.getRange(i+2,2).setValue(prenom_kholleur);
    sheet.getRange(i+2,3).setValue(matiere);
  }
  sheet.getRange(2,4,les_kholleurs.length-1,5).setDataValidation(SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInList(["Oui","Non"], true)
        .build());
  sheet.getRange(1,1,les_kholleurs.length,8).setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange(1,1,les_kholleurs.length,8).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  var banding = sheet.getRange(1,1,les_kholleurs.length,8).getBandings()[0];
  banding.setHeaderRowColor('#5b95f9')
  .setFirstRowColor('#ffffff')
  .setSecondRowColor('#e8f0fe')
  .setFooterRowColor(null);
  sheet.getRange(1,1,1,8).setFontWeight('bold').setFontSize(11);
  sheet.autoResizeColumns(1, 8);
}

function retourne_matiere(matiere)
{
  if (les_matieres.length==0)
    charge_matieres();
  for (let i = 0; i < les_matieres.length;i++)
    if (les_matieres[i]["nom"] == matiere)
      return les_matieres[i];
  msg_log = "Erreur retourne_matiere : matiere inconnue : " + matiere.
  console.error(msg_log);
  ecrit_log(LOG_ERROR,"retourne_eleve",msg_log);

}

function retourne_eleve(nom,prenom)
{
  if (les_eleves.length==0)
    charge_eleves();
  for (let i = 0; i < les_eleves.length; i++)
  {
    if (les_eleves[i]["nom"] == nom && les_eleves[i]["prenom"] == prenom)
      return les_eleves[i];
  }
  msg_log = "Erreur retourne_eleve : élève non trouvé : " + nom + " " + prenom;
  console.error(msg_log);
  ecrit_log(LOG_ERROR,"retourne_eleve",msg_log);
  return -1;
}

function retourne_kholleur(nom,prenom)
{
  if (les_kholleurs.length==0)
    charge_kholleurs();
  for (let i = 0; i < les_kholleurs.length; i++)
  {
    if (les_kholleurs[i]["nom"] == nom && les_kholleurs[i]["prenom"] == prenom)
      return les_kholleurs[i];
  }
  msg_log = "Erreur retourne_kholleur : kholleur non trouvé : " + nom + " " + prenom;
  console.error(msg_log);
  ecrit_log(LOG_ERROR,"retourne_kholleur",msg_log);
  return -1;
}

function retourne_professeur(nom,prenom)
{
  if (les_professeurs.length == 0)
    charge_professeurs();
  for (let i = 0; i < les_professeurs.length; i++)
  {
    if (les_professeurs[i]["nom"] == nom && les_professeurs[i]["prénom"] == prenom)
      return les_professeurs[i];
  }
  msg_log = "Erreur retourne_professeur : professeur non trouvé : " + nom + " " + prenom;
  console.error(msg_log);
  ecrit_log(LOG_ERROR,"retourne_professeur",msg_log);
  return -1;
}