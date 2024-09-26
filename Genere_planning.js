const APRES_CHAQUE_PASSAGE ="5 minutes après chaque passage";
const TOUS_LES_4_PASSAGES = "10 minutes tous les 4 passages";
const PAS_DE_PAUSE = "Pas de pause";
var html;

function createDropdown(cell, list) {
  var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(list) // Set 'false' for no checkbox
      .build();
  cell.setDataValidation(rule);
}

function test_button()
{
  /*
 var htmlOutput = HtmlService.createHtmlOutput(`<p>hello world</p>`);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "My Popup");
  */
   showSidebar();

}

function affiche_planning_interrogations()
{
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_PLANNING_KHOLLES).activate();
}

function choisit_planning_interrogations()
{
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_PLANNING_KHOLLES).activate();
  choisit_planning_kholle();
}

function choisit_planning_kholle() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('liste_plannings').setTitle("Choix d'un planning de khôlle");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function getAllFilledLines() {
  const donnees = sheetToAssociativeArray(FEUILLE_PLANNING_KHOLLES);
  return donnees.filter(row=>row["Date"] instanceof Date && row["Date"] != "");
}

function construit_liste_plannings() {
  // Get data from spreadsheet cells
  let html = "";
  let lines = getAllFilledLines();
  /*lines.sort(function(l1, l2) 
      {
        const d1 = new Date(l1["Date"]);
        const d2 = new Date(l2["Date"]);
        if (d1<d2) 
          return 1;
        else
          return -1;
    });*/
  let css = "";
  let deja_checked = false;
  for (let i = 0;i < lines.length;i++)
  {
    const date =  formatDate(lines[i]["Date"]);
    const planning_enregistre = lines[i]["Planning enregistré"];
    const annee = lines[i]["Année"];
    const discipline = lines[i]["Discipline"];
    let checked;
    if (planning_enregistre=="Oui")
    {
      checked = '';
      css="color:red;";
    }
    else
    {
      if (!deja_checked)
      {
        checked = "checked='checked'";
        deja_checked = true;
      }
      css = "color:blue;"
    }
    html += "<input type='radio' name='rd_dates_planning' " + checked + " value='"+i+"'><label style="+css +">" + date + " - " + discipline + " - " + annee +"</label><br>";
    //alert(html);
  }
  return html;
}

function recupere_message()
{
  //console.log("hmmman");
  /*let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CALCUL_TEMPORAIRES);
  sheet.getRange(1,8).setValue("hmmm");*/
  const msg = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").getValue();
  return msg;
}

function recupere_id_annuaire()
{
  let html = "<div>Consulter / mettre à jour l'annuaire : </div>";
  html += "<ul><li><p>des examinateurs (avec leurs disponibilités)</p></li>";
  html += "<li><p>des professeurs référents</p></li>";
  html += "<li><p>des étudiants (L1 et L2)</p></li>";
  html += "</ul><BR><BR>";
  const lien = "https://docs.google.com/spreadsheets/d/" + ID_ANNUAIRE + "/edit";
  //html += "<div>>Annuaire 2024-2025</a></div>";
  //html += "<BR><center><button type=\"button\" class=\"button\" onclick=\"location.href = '" + //lien + "'\">Annuaire 2024-2025</button></center>";
  //let button = "<BR><center><button type=\"button\" class=\"button\" onclick=\"location.href ='" + lien +"'\">Annuaire 2024-2025</button></center>";
  let button = "<center><button type=\"button\" class=\"button\" onclick=\"window.open('" + lien +"','_blank');\">Annuaire 2024-2025</button></center>";
  html += button;
  return html;
}

function hm()
{
  //console.log("hmmman");
}

function genere_p()
{
  choisit_le_planning(2);
}

function construit_nom_onglet_planning(annee,date_kholle,discipline)
{
  return "Planning Khôlle : " + annee + " - " + date_kholle + " - " + discipline;
}

function choisit_le_planning(ind)
{
  const line = getAllFilledLines()[ind];
  const annee = line["Année"];
  const date_kholle = formatDate(line["Date"]);
  const discipline = line["Discipline"];
  const nom_onglet = construit_nom_onglet_planning(annee,date_kholle,discipline);
  const sh_temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CALCUL_TEMPORAIRES);
  const nom_onglet_precedent = sh_temp.getRange(1,1).getValue();
  if (nom_onglet_precedent.startsWith("Planning Khôlle :") && nom_onglet != nom_onglet_precedent)
    masque_autre_onglet_visible(nom_onglet_precedent);
  sh_temp.getRange(1,1).setValue(nom_onglet);
  SpreadsheetApp.getActiveSpreadsheet().setNamedRange('nom_onglet_planning',  sh_temp.getRange(1,1));
  if (!positionne_onglet(nom_onglet)) 
    positionne_onglet_choix_kholleurs(nom_onglet);
  else
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet).activate();
}

function cherche_si_kholle_meme_semaine_meme_discipline(annee,discipline,date_kholle)
{
  const sh_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_PLANNING_KHOLLES);
  const no_semaine = getWeekNumber(date_kholle);
  let donnees = sh_planning.getDataRange().getValues();
  donnees = donnees.filter(
                      function(row) 
                                  {
                                    const autre_semaine = getWeekNumber(row[1]);
                                    const condition = row[0] == annee && autre_semaine == no_semaine && row[2] == discipline && row[3] == "Oui" && date_kholle != row[1];
                                    //console.log(condition);
                                    return condition;
                                  });
  if (donnees.length>0)
    return donnees[0][1];
  return null;
}

function retourne_eleves_deja_kholles_meme_semaine_meme_discipline(annee,date_kholle,discipline)
{
  const sh_histo_eleves = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_HISTO_ELEVES);
  const no_semaine = getWeekNumber(date_kholle);
  let donnees = sh_histo_eleves.getDataRange().getValues();
  donnees = donnees.filter(row => row[2] == annee && getWeekNumber(row[6]) == no_semaine && row[9] == discipline);
  if (donnees.length==0)
  {
    msg_log = "Erreur pas d'éléves trouvés la même semaine, date kholle : " + date_kholle + " classe : " + annee + " discipline : " + discipline;
    ecrit_log(LOG_WARNING,"retourne_eleves_deja_kholles_meme_semaine_meme_discipline",msg_log);
    return [];
  }
  let les_eleves = [];
  for (donnee of donnees)
  {
    les_eleves.push({"prénom":donnee[1],"nom":donnee[0]});
  }
  return les_eleves;
}

function positionne_onglet_choix_kholleurs(nom_onglet)
{
   const sh_choix_kholleurs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CHOIX_KHOLLEURS);
   sh_choix_kholleurs.activate();
   sh_choix_kholleurs.clearContents();
   sh_choix_kholleurs.getRange("A:B").clearDataValidations();
   sh_choix_kholleurs.getRange(8,2,6,2).setBackground(null);
   const infos_planning = retourne_infos_from_nom_onglet_planning(nom_onglet);
   const annee = infos_planning["annee"];
   let date_kholle = infos_planning["date_kholle"];
   const date_split = date_kholle.split("/");
   date_kholle = new Date(date_split[2],date_split[1]-1,date_split[0]);
   const discipline = infos_planning["discipline"];
   const infos = cherche_si_kholle_meme_semaine_meme_discipline(annee,discipline,date_kholle);
   let eleves_deja_kholles_meme_semaine_meme_discipline;
   if (infos == null)
    eleves_deja_kholles_meme_semaine_meme_discipline = [];
   else
    eleves_deja_kholles_meme_semaine_meme_discipline = retourne_eleves_deja_kholles_meme_semaine_meme_discipline(annee,infos,discipline);
   const eleves_annee_discipline = charge_eleves_annee_discipline(annee,discipline,eleves_deja_kholles_meme_semaine_meme_discipline);
   const nb_eleves = eleves_annee_discipline.length;
   let donnees = [];
   const date_complete = date_kholle.getDayName() + " " + date_kholle.getDate() + " " + date_kholle.getMonthName() + " " + date_kholle.getFullYear();
   const titre = "Paramètres de la khôlle en " + annee + " du " + date_complete + " - " + discipline;
   donnees.push([titre,"",""]);
   donnees.push(["Nombre d'étudiants :",nb_eleves,nb_eleves]);
   const nb_kholleurs = Math.max(charge_kholleurs_discipline(annee,date_kholle,discipline,true).length,MAX_KHOLLEURS_PAR_SEANCE);
   donnees.push(["Nombre d'examinateurs :",nb_kholleurs,""]);
   donnees.push(["Durée interrogation (en minutes):",20,""]);
   let preparation;
   if (est_LV(discipline))
    preparation = 20;
  else
    preparation = 30;
   donnees.push(["Durée préparation (en minutes):",preparation,""]);
   let pause;
   if (MODE_SIMULATION)
    pause = shuffle(["5 min après chaque passage","10 min tous les 3 passages","Pas de pause"])[0];
   donnees.push(["Pause :",pause,""]);
   donnees.push(["Horaire première khôlle :","18:10",""]);
   sh_choix_kholleurs.getRange(1,1,7,3).setValues(donnees);
   sh_choix_kholleurs.getRange('B3').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .setHelpText('Saisissez un nombre compris entre 1 et ' + nb_kholleurs)
  .requireNumberBetween(1, nb_kholleurs)
  .build());
   const pauses = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("Pauses_kholles").getValues();
   sh_choix_kholleurs.getRange('B6').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireValueInList(pauses, true)
  .build());
  const sh_temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CALCUL_TEMPORAIRES);
  let donnees_temp = [];
  donnees_temp.push(["Date kholle",date_kholle,""]);
  donnees_temp.push(["Discipline",discipline,""]);
  donnees_temp.push(["Année",annee,""]);
  sh_temp.getRange(5,1,3,3).setValues(donnees_temp);
  sh_temp.getRange("C2").setValue(ETAT_KHOLLEURS_NON_PROPOSES);
}

function proposer_kholleurs_avec_dispo()
{
    proposer_kholleurs(true);
}  

function proposer_kholleurs_sans_dispo()
{
    proposer_kholleurs(false);
}  


function proposer_kholleurs(avec_dispo)
{
  const sh_temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CALCUL_TEMPORAIRES);
  const donnees_kholle = sh_temp.getRange(5,1,3,2).getValues();
  const date_kholle = donnees_kholle[0][1];
  const discipline = donnees_kholle[1][1];
  const annee = donnees_kholle[2][1];
  const sh_choix_kholleurs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CHOIX_KHOLLEURS);
  const donnees_params_kholle = sh_choix_kholleurs.getRange(2,2,6,1).getValues();
  const nb_eleves = donnees_params_kholle[0][0];
  const nb_examinateurs = donnees_params_kholle[1][0];
  const duree_kholle = donnees_params_kholle[2][0];
  const duree_preparation = donnees_params_kholle[3][0];
  const pause = donnees_params_kholle[4][0];
  const heure_premiere_kholle = donnees_params_kholle[5][0];
  const hh = new Date(heure_premiere_kholle).getHours();
  const mm = new Date(heure_premiere_kholle).getMinutes();
  let donnees = [];
  let kholleurs_discipline = charge_kholleurs_discipline(annee,date_kholle,discipline,avec_dispo);
  const nb_kholleurs = kholleurs_discipline.length;
  if (nb_kholleurs == 0)
  {
     SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue("Aucun examinateur n'est disponible ce jour là !");
    showSideBarInfos("Paramètres de la khôlle");
    return;
  }
   sh_choix_kholleurs.getRange('B3').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .setHelpText('Saisissez un nombre compris entre 1 et ' + nb_kholleurs)
  .requireNumberBetween(1, nb_kholleurs)
  .build());
  const ctrl = controle_saisies_parametres_kholle(nb_examinateurs,duree_kholle,pause,duree_preparation,heure_premiere_kholle,nb_kholleurs);
  if (ctrl!=true)
  {
     SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue(ctrl);
    showSideBarInfos("Paramètres de la khôlle");
    return;
  }
  donnees.push(["Examinateurs","Disponibilités","Nombre d'étudiants interrogés","Horaire première khôlle","Fin dernière khôlle"]);
  
  //kholleurs_discipline = shuffle(kholleurs_discipline);
  kholleurs_discipline = retourne_kholleurs_tries(kholleurs_discipline);
  let les_kholleurs_affiches = [];
  for (let i = 0; i < kholleurs_discipline.length;i++)
  les_kholleurs_affiches.push(kholleurs_discipline[i]["civilite"] + " " + kholleurs_discipline[i]["nom"] + " " + kholleurs_discipline[i]["prenom"]);
  const eleves_annee_discipline = charge_eleves_annee_discipline(annee,discipline);
  sh_choix_kholleurs.getRange(2,2).setValue(nb_eleves);
  const repartition_nb_eleves_par_kholleurs = determine_repartition_nb_eleves_par_kholleurs(nb_eleves,nb_examinateurs);
  const duree_pause_apres_chaque_passage = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("duree_pause_apres_chaque_passage").getValue();
  const duree_pause_passages = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("duree_pause_passages").getValue();
  const nb_passages_entre_deux_pauses
  = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nb_passages_entre_deux_pauses").getValue();
  sh_choix_kholleurs.getRange("A:A").clearDataValidations();  
  for (let ind_examinateur = 1; ind_examinateur <= nb_examinateurs; ind_examinateur++)
  {
    sh_choix_kholleurs.getRange(ind_examinateur+8,1).setDataValidation(SpreadsheetApp.newDataValidation()
      .setAllowInvalid(true)
      .requireValueInList(les_kholleurs_affiches, true)
      .build());
    const nb_etudiants_kholleur = repartition_nb_eleves_par_kholleurs[ind_examinateur-1];
    let heure = new Date(date_kholle);
    const plage_dispo = kholleurs_discipline[ind_examinateur-1]["disponibilités"];
    const heure_split = plage_dispo.split(" - ");
    const heure_debut = heure_split[0];
    const hh = parseInt(heure_debut.substring(0,2));
    const mm = parseInt(heure_debut.substring(3,5));
    heure.setHours(hh);
    heure.setMinutes(mm);
    const heure_fin_kholle = calcule_heure_fin_kholle(heure,duree_kholle,pause,duree_pause_apres_chaque_passage,duree_pause_passages,nb_passages_entre_deux_pauses,nb_etudiants_kholleur);
    sh_choix_kholleurs.getRange(8,4,8+nb_etudiants_kholleur,2).setNumberFormat('hh":"mm');
    donnees.push([les_kholleurs_affiches[ind_examinateur-1],plage_dispo,nb_etudiants_kholleur,heure_debut,heure_fin_kholle]);
  }
  for (let ind_examinateur = nb_examinateurs + 1 ; ind_examinateur <= 6; ind_examinateur++)
    donnees.push(["","","","",""]);
  sh_choix_kholleurs.getRange(8,1,7,5).setValues(donnees);
  sh_choix_kholleurs.getRange("A8:E8").setBackground('#e8c5eb');
  sh_choix_kholleurs.getRange(9,2,nb_examinateurs,2).setBackground('#d9ead3');
  sh_choix_kholleurs.getRange(9+nb_examinateurs,2,8-nb_examinateurs,3).setBackground(null);
  sh_temp.getRange("C2").setValue(ETAT_KHOLLEURS_PROPOSES);
}


function masque_autre_onglet_visible(nom_onglet)
{
  try
  {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet).hideSheet();
  }
  catch (err)
  {

  }
}

function masque_autres_plannings(annee,date_kholle,discipline)
{
  for (const onglet of SpreadsheetApp.getActiveSpreadsheet().getSheets())
  {
    const autre_onglet = onglet.getName();
    const autre_onglet_split = autre_onglet.split(" : ");
    if (autre_onglet_split.length != 0)
    {
      if (autre_onglet_split[0] == "Planning Khôlle")
      {
        const autre_onglet_split2 = autre_onglet_split[1];
        const annee_autre_onglet = autre_onglet_split2[0];
        const date_kholle_autre_onglet = autre_onglet_split2[1];
        const discipline_autre_onglet = autre_onglet_split2[2];
        if (annee != annee_autre_onglet || date_kholle_autre_onglet != date_kholle || discipline != discipline_autre_onglet)
          onglet.hideSheet();
      }
    }
  }
}

function positionne_onglet(nom_onglet)
{
  let existingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet);
  return existingSheet;
}

function ecrire_planning_avec_kholleurs_choisis()
{
 
  const ctrl = controle_saisies_kholleurs_choisis();
  if (ctrl!=true)
  {
     SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue(ctrl);
    showSideBarInfos("Paramètres de la khôlle");
    return;
  }
  ecrit_planning_kholle_V2();
 
}

function controle_saisies_parametres_kholle(nb_examinateurs,duree_kholle,pause,duree_preparation,heure_premiere_kholle,nb_kholleurs_max)
{
  if (!isInteger(nb_examinateurs))
    return "Saisir un nombre d'examinateurs compris entre 1 et " + nb_kholleurs_max;
  if (nb_examinateurs < 1 || nb_examinateurs > nb_kholleurs_max)  
    return "Saisir un nombre d'examinateurs compris entre 1 et " + nb_kholleurs_max;
  if (!isInteger(duree_kholle))
    return "Saisir une durée d'interrogation (en minutes) valide.";
  if (!isInteger(duree_preparation))
    return "Saisir une durée de préparation (en minutes) valide.";
  if (pause == "")
    return "Saisir un mode de pause proposé dans la liste déroulante.";
  if (!isValidHourMinute(heure_premiere_kholle))
    return "Saisir l'horaire de la première khôlle au format hh:mm";
  return true;
}

function controle_saisies_kholleurs_choisis()
{
   const sh_temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CALCUL_TEMPORAIRES);
  const etat_planning = sh_temp.getRange("C2").getValue();
  if (etat_planning < ETAT_KHOLLEURS_PROPOSES)
    return "Examinateurs non encore proposés";
   const donnees_kholle = sh_temp.getRange(5,1,3,2).getValues();
   const discipline = donnees_kholle[1][1];
   const annee = donnees_kholle[2][1];
   const sh_choix_kholleurs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CHOIX_KHOLLEURS);
   const donnees = sh_choix_kholleurs.getDataRange().getValues();
   const nb_eleves =  donnees[1][1];
   const nb_examinateurs = donnees[2][1];
    if (!isInteger(nb_examinateurs))
  {
     return "Saisir le nombre d'examinateurs";
  }
  const eleves_annee_discipline = charge_eleves_annee_discipline(annee,discipline);
  let nb_eleves_saisis = 0;
  const duree_interrogation = donnees[3][1];
  if (!isInteger(duree_interrogation))
  {
     return "Saisir la durée de l'interrogation";
  }
  let les_kholleurs = new Set();
  for (let ind_examinateur = 1; ind_examinateur <= nb_examinateurs; ind_examinateur++)
  {
    const kholleur =donnees[7+ind_examinateur][0];
    if ( kholleur == "")
    {
      return "Au moins un examinateur non saisi";
    }
    else
    {
      if (les_kholleurs.has(kholleur))
      {
        return "2 fois le même examinateur saisi";
      }
      else
        les_kholleurs.add(kholleur);
    }
   const nb_etudiants_kholleur =donnees[7+ind_examinateur][2];
   if (!isInteger(nb_etudiants_kholleur))
   {
      return "Saisir un nombre d'étudiants pour chaque examinateur";
   }
   else
    nb_eleves_saisis += nb_etudiants_kholleur;
   const heure_debut_kholle =donnees[7+ind_examinateur][3];
   if (!isValidHourMinute(heure_debut_kholle))
   {
    return "Saisir une heure de début de khôlle valide (format hh:mm) pour chaque examinateur";
   }
  }
  if (nb_eleves != nb_eleves_saisis)
  {
    return "Le nombre d'étudiants à interroger : " + nb_eleves + " ne correspond pas au nombre total d'étudiants assignés " + nb_eleves_saisis + " aux " + nb_examinateurs + " examinateurs";
  }
  return true;
}

function ecrit_planning_kholle_V2()
{
   const sh_temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CALCUL_TEMPORAIRES);
   const donnees_kholle = sh_temp.getRange(5,1,3,2).getValues();
   const date_kholle = donnees_kholle[0][1];
   const discipline = donnees_kholle[1][1];
   const annee = donnees_kholle[2][1];
   const sh_choix_kholleurs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CHOIX_KHOLLEURS);
   const donnees = sh_choix_kholleurs.getDataRange().getValues();
   const nom_onglet_prec = sh_temp.getRange(1,1).getValue();
   const nom_onglet =  construit_nom_onglet_planning(annee,formatDate(date_kholle),discipline);
   sh_temp.getRange(1,1).setValue(nom_onglet);
   const nb_eleves_a_choisir = donnees[1][1];
   const nb_examinateurs = donnees[2][1];
   const duree_kholle = donnees[3][1];
   const duree_preparation = donnees[4][1];
   const pause = donnees[5][1];
   const old_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_prec);
   if (old_sheet)
      old_sheet.hideSheet();
   createOrReplaceSheet(nom_onglet);
   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet);
   sheet.activate();
   let donnees_planning = [];
   const titre = "khôlles - " + annee + " - " + discipline + " - " + formatDate(date_kholle);
   let ligne = [titre,"","","",""];
   donnees_planning.push(ligne);
  let no_ligne = 3;
   let no_ligne_debut = 3;
   const duree_pause_apres_chaque_passage = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("duree_pause_apres_chaque_passage").getValue();
   const duree_pause_passages = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("duree_pause_passages").getValue();
   const nb_passages_entre_deux_pauses
    = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nb_passages_entre_deux_pauses").getValue();
   let libelle_pause;
   if (pause.includes("tous"))
   {
    libelle_pause = "Pause : " + duree_pause_passages + " min";
   }
   else
   {
    if (pause.includes("chaque"))
    {
      libelle_pause = "Pause : " + duree_pause_apres_chaque_passage + "min";
    }
   }
   ligne = ["Salle","Examinateurs","Candidats","Préparation","Passage"];
   donnees_planning.push(ligne);
   sheet.getRange(1, 7).setValue("Professeur référent");
   charge_matieres();
   let kholleurs_discipline = charge_kholleurs_discipline(annee,date_kholle,discipline,false);
   const professeurs_discipline = charge_professeurs_discipline(annee,discipline);
   let professeurs_referents = [];
   for (let i = 0; i < professeurs_discipline.length; i++)
   {
    professeurs_referents.push(professeurs_discipline[i][4] + " " + professeurs_discipline[i][2] + " " + professeurs_discipline[i][3]);
   }
   sheet.getRange(1,8).setDataValidation(SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInList(professeurs_referents, true)
        .build());
   if (professeurs_referents.length > 0)
    sheet.getRange(1,8).setValue(professeurs_referents[0]);     
   //kholleurs_discipline = shuffle(kholleurs_discipline);
   kholleurs_discipline = retourne_kholleurs_tries(kholleurs_discipline);
   charge_salles_kholles();
  const infos = cherche_si_kholle_meme_semaine_meme_discipline(annee,discipline,date_kholle);
  let eleves_deja_kholles_meme_semaine_meme_discipline;
  if (infos == null)
    eleves_deja_kholles_meme_semaine_meme_discipline = [];
  else
    eleves_deja_kholles_meme_semaine_meme_discipline = retourne_eleves_deja_kholles_meme_semaine_meme_discipline(annee,infos,discipline);
  const eleves_annee_discipline = charge_eleves_annee_discipline(annee,discipline,eleves_deja_kholles_meme_semaine_meme_discipline);
   const nb_eleves = eleves_annee_discipline.length;
   let repartition_nb_eleves_par_kholleur = [];
   let heure_debut_kholle_par_kholleur = [];
   let les_examinateurs_choisis = [];
  for (let i = 0;i < nb_examinateurs;i ++)
  {
    repartition_nb_eleves_par_kholleur.push(donnees[8+i][2]);
    heure_debut_kholle_par_kholleur.push(donnees[8+i][3]);
  } 
   for (let ind_examinateur = 1; ind_examinateur <= nb_examinateurs; ind_examinateur++)
   {
      const kholleur = donnees[7+ind_examinateur][0];
      const kholleur_split = kholleur.split(" ");
      const civilite_kholleur = kholleur_split[0];
      const nom_kholleur = kholleur_split[1];
      const prenom_kholleur = kholleur_split[2];
      const civilite_nom_prenom_kholleur = civilite_kholleur + " " + nom_kholleur + " " + prenom_kholleur;
      const civilite_nom_kholleur = civilite_kholleur + " " + nom_kholleur;
      les_examinateurs_choisis.push(civilite_nom_prenom_kholleur);
      no_ligne_debut = no_ligne;
      sheet.getRange(no_ligne,1).setDataValidation(SpreadsheetApp.newDataValidation()
        .setAllowInvalid(true)
        .requireValueInList(les_salles_kholles, true)
        .build());
      let hh = new Date(heure_debut_kholle_par_kholleur[ind_examinateur-1]).getHours();
      let mm = new Date(heure_debut_kholle_par_kholleur[ind_examinateur-1]).getMinutes();
      let heure = new Date(date_kholle);
      heure.setHours(hh);
      heure.setMinutes(mm);
      const nb_etudiants_kholleur = repartition_nb_eleves_par_kholleur[ind_examinateur-1];
      for (let ind_passage = 0; ind_passage < nb_etudiants_kholleur; ind_passage++)
      {
        if (((ind_passage % nb_passages_entre_deux_pauses == 0 && pause.includes("tous")) || pause.includes("chaque")) && ind_passage != 0)
        {
          ligne = ["","",libelle_pause,"",""];
          sheet.getRange(no_ligne,3,1,3).merge();
          sheet.getRange(no_ligne_debut,1).setHorizontalAlignment('center');
          no_ligne++;
          donnees_planning.push(ligne);
        }
          ligne = [""];
          if (ind_passage == 0)
            ligne.push(civilite_nom_kholleur);
          else
            ligne.push("");
          ligne.push("");
          sheet.getRange(no_ligne,3).setDataValidation(SpreadsheetApp.newDataValidation()
          .setAllowInvalid(true)
          .requireValueInList(eleves_annee_discipline, true)
          .build());
          let creneau_horaire = formate_creneau_horaire(heure,duree_kholle);
          let heure_debut_preparation = calcule_heure_debut_preparation(heure,duree_preparation);
          ligne.push(heure_debut_preparation);
          ligne.push(creneau_horaire);
          heure = heure_suivante(heure,duree_kholle,ind_passage,pause,duree_pause_apres_chaque_passage,duree_pause_passages,nb_passages_entre_deux_pauses,nb_etudiants_kholleur);
        no_ligne++;
        donnees_planning.push(ligne);
      }
      sheet.getRange(no_ligne_debut,1,no_ligne-no_ligne_debut,1).setHorizontalAlignment('center');
      sheet.getRange(no_ligne_debut,1,no_ligne-no_ligne_debut,1).mergeVertically();
      sheet.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES+ind_examinateur).setValue(no_ligne_debut);
      sheet.getRange(LIGNE_MEMO+1,COLONNE_MEMO_NB_ELEVES+ind_examinateur).setValue(no_ligne-1);
      sheet.getRange(no_ligne_debut,1).setVerticalAlignment('middle');
      sheet.getRange(no_ligne_debut,2,no_ligne-no_ligne_debut,1).mergeVertically();
      sheet.getRange(no_ligne_debut,2).setVerticalAlignment('middle');
   }
   sheet.getRange(1,1,no_ligne-1,5).setValues(donnees_planning);
   entete_planning_kholle();
   mef_tableau(sheet,no_ligne - 1,5);
   sheet.setColumnWidth(1, 80);
   sheet.getRange(4,5).setHorizontalAlignment('center');
   sheet.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_EXAMINATEURS).setValue(nb_examinateurs);
   sheet.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES).setValue(nb_eleves);
   sheet.getRange(LIGNE_MEMO+1,COLONNE_MEMO_NB_ELEVES).setValue(nb_eleves_a_choisir);
   let donnees_repartitions_eleves = [];
   for (let i = 0;i < nb_examinateurs;i++)
   {
    donnees_repartitions_eleves.push([les_examinateurs_choisis[i],repartition_nb_eleves_par_kholleur[i]]);
   }
   sheet.getRange(LIGNE_MEMO + 10,COLONNE_MEMO_NB_EXAMINATEURS,nb_examinateurs,2).setValues(donnees_repartitions_eleves);
  sh_temp.getRange("C2").setValue(PLANNING_ECRIT_NON_ENREGISTRE);
}

function ecrit_planning_kholle(nom_onglet,line)
{
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet);
   sheet.clearContents();
   let annee = line[0];
   let date_kholle = line[1].getDayName() + " " + line[1].getDate() + " " + line[1].getMonthName() + " " + line[1].getFullYear();
   let discipline = line[2];
   let duree_kholle = line[3];
   sheet.getRange(1,1).setValue("khôlles - " + annee + " - " + discipline + " - " + date_kholle);
   entete_planning_kholle();
   let pause = line[4];
   let nb_examinateurs = line[5];
   let no_ligne = 3;
   let no_ligne_debut = 3;
   const duree_pause_apres_chaque_passage = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("duree_pause_apres_chaque_passage").getValue();
   const duree_pause_passages = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("duree_pause_passages").getValue();
   const  nb_passages_entre_deux_pauses
    = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nb_passages_entre_deux_pauses").getValue();
   sheet.getRange(2, 1).setValue("Salle");
   sheet.getRange(2, 2).setValue("Examinateurs");
   sheet.getRange(2, 3).setValue("Candidats");
   sheet.getRange(2, 4).setValue("Préparation");
   sheet.getRange(2, 5).setValue("Passage");
   sheet.getRange(1, 7).setValue("Professeur référent");
   charge_matieres();
   let kholleurs_discipline = charge_kholleurs_discipline(annee,line[1],discipline,false);
   const professeurs_discipline = charge_professeurs_discipline(annee,discipline);
   let professeurs_referents = [];
   for (let i = 0; i < professeurs_discipline.length; i++)
   {
    professeurs_referents.push(professeurs_discipline[i][4] + " " + professeurs_discipline[i][2] + " " + professeurs_discipline[i][3]);
   }
   sheet.getRange(1,8).setDataValidation(SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInList(professeurs_referents, true)
        .build());
   if (professeurs_referents.length > 0)
    sheet.getRange(1,8).setValue(professeurs_referents[0]);     
   //kholleurs_discipline = shuffle(kholleurs_discipline);
   kholleurs_discipline = retourne_kholleurs_tries(kholleurs_discipline);
   let les_kholleurs_affiches = [];
   for (let i = 0; i < kholleurs_discipline.length;i++)
    les_kholleurs_affiches.push(formate_nom_kholleur(kholleurs_discipline[i]["civilite"],kholleurs_discipline[i]["prenom"],kholleurs_discipline[i]["nom"]));
   let duree_preparation = lit_preparation(discipline);
   charge_salles_kholles();
   let eleves_annee_discipline = charge_eleves_annee_discipline(annee,discipline);
   let nb_eleves = eleves_annee_discipline.length;
    const repartition_nb_eleves_par_kholleurs = determine_repartition_nb_eleves_par_kholleurs(nb_eleves,nb_examinateurs);
   let no_passage;
   for (let ind_examinateur = 1; ind_examinateur <= nb_examinateurs; ind_examinateur++)
   {
      no_ligne_debut = no_ligne;
      sheet.getRange(no_ligne,1).setDataValidation(SpreadsheetApp.newDataValidation()
        .setAllowInvalid(true)
        .requireValueInList(les_salles_kholles, true)
        .build());
      sheet.getRange(no_ligne,2).setDataValidation(SpreadsheetApp.newDataValidation()
        .setAllowInvalid(true)
        .requireValueInList(les_kholleurs_affiches, true)
        .build());
      let hh = new Date(line[6]).getHours();
      let mm = new Date(line[6]).getMinutes();
      let heure = new Date(line[1]);
      heure.setHours(hh);
      heure.setMinutes(mm);
      no_passage = 0;
      const nb_eleves_kholleur = repartition_nb_eleves_par_kholleurs[ind_examinateur-1];
      for (let ind_passage = 0; ind_passage < nb_eleves_kholleurs ; ind_passage++)
      {
        if (ind_passage % nb_passages_entre_deux_pauses == 0 && pause.includes("tous") && no_passage != 0)
        {
          sheet.getRange(no_ligne,3).setValue("Pause : " + duree_pause_apres_chaque_passage + "min");
          sheet.getRange(no_ligne,3,1,3).merge();
          sheet.getRange(no_ligne_debut,1).setHorizontalAlignment('center');
          no_ligne++;
        }
        sheet.getRange(no_ligne,3).setDataValidation(SpreadsheetApp.newDataValidation()
        .setAllowInvalid(true)
        .requireValueInList(eleves_annee_discipline, true)
        .build());
        let creneau_horaire = formate_creneau_horaire(heure,duree_kholle);
        let heure_debut_preparation = calcule_heure_debut_preparation(heure,duree_preparation);
        sheet.getRange(no_ligne,4).setValue(heure_debut_preparation);
        sheet.getRange(no_ligne,5).setValue(creneau_horaire);
        heure = heure_suivante(heure,duree_kholle,ind_passage,pause,duree_pause_apres_chaque_passage,duree_pause_passages,nb_passages_entre_deux_pauses,nb_eleves_kholleur);
        no_ligne++;
        no_passage++;
      }
      sheet.getRange(no_ligne_debut,1,no_ligne-no_ligne_debut,1).setHorizontalAlignment('center');
      sheet.getRange(no_ligne_debut,1,no_ligne-no_ligne_debut,1).mergeVertically();
      sheet.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES+ind_examinateur).setValue(no_ligne_debut);
      sheet.getRange(LIGNE_MEMO+1,COLONNE_MEMO_NB_ELEVES+ind_examinateur).setValue(no_ligne-1);
      sheet.getRange(no_ligne_debut,1).setVerticalAlignment('middle');
      sheet.getRange(no_ligne_debut,2,no_ligne-no_ligne_debut,1).mergeVertically();
      sheet.getRange(no_ligne_debut,2).setVerticalAlignment('middle');
   }
   mef_tableau(sheet,no_ligne - 1,5);
   sheet.autoResizeColumns(7, 8);
   sheet.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_EXAMINATEURS).setValue(nb_examinateurs);
   sheet.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES).setValue(nb_eleves);
}

function mef_tableau(sheet,no_ligne,no_colonne)
{
  //console.log(sheet.getName());
  //console.log(sheet.getRange(1,1).getValue());
  try
  {
    sheet.getRange(1,1,1,no_colonne).setFontWeight('bold');
    sheet.getRange(1,1,no_ligne,no_colonne).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(1,1,no_ligne,no_colonne).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    var banding = sheet.getRange(1,1,no_ligne,no_colonne).getBandings()[0];
    banding.setHeaderRowColor('#5b95f9')
    .setFirstRowColor('#ffffff')
    .setSecondRowColor('#e8f0fe')
    .setFooterRowColor(null);
    sheet.getRange(1,1,no_ligne,no_colonne).setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
    sheet.getRange(2,1,1,no_colonne).setFontWeight('bold').setFontSize(11);
  }
  catch (err)
  {

  }
  sheet.getRange(1,1,1,no_colonne).moveTo(sheet.getRange(100,100));
  sheet.autoResizeColumns(1, no_colonne);
  sheet.getRange(100,100,1,no_colonne).moveTo(sheet.getRange(1,1));
  for (let i = 2; i <= no_colonne ; i++)
     sheet.setColumnWidth(i,sheet.getColumnWidth(i)*1.2);
}

function ajoute_duree_heure(heure,duree)
{
  let h_suivante = new Date();
  h_suivante.setHours(heure.getHours());
  h_suivante.setMinutes(heure.getMinutes()+duree);
  return h_suivante;
}

function heure_suivante(heure,duree_kholle,ind_passage,pause,duree_pause_apres_chaque_passage,duree_pause_passages,nb_passages_entre_deux_pauses,nb_etudiants_kholleur)
{
  let h_suivante = new Date();
  const hh = heure.getHours();
  const mn = heure.getMinutes();
  h_suivante.setHours(hh)
  h_suivante.setMinutes(mn);
  if (pause.includes("Pas"))
    h_suivante.setMinutes(heure.getMinutes()+duree_kholle);
  else
  {
    if (pause.includes("après"))
    {
      if (ind_passage != nb_etudiants_kholleur - 1)
        h_suivante.setMinutes(heure.getMinutes()+duree_kholle+duree_pause_apres_chaque_passage);
      else
      h_suivante.setMinutes(heure.getMinutes()+duree_kholle);

    }
    else
    {
      if ((ind_passage + 1) % nb_passages_entre_deux_pauses == 0 && ind_passage != (nb_etudiants_kholleur - 1))
          h_suivante.setMinutes(heure.getMinutes()+duree_kholle+duree_pause_passages);
        else
          h_suivante.setMinutes(heure.getMinutes()+duree_kholle);
    }
  }
  return h_suivante;
}

function calcule_heure_fin_kholle(heure,duree_kholle,pause,duree_pause_apres_chaque_passage,duree_pause_passages,nb_passages_entre_deux_pauses,nb_etudiants_kholleur)
{
  for (let ind_passage = 0;ind_passage < nb_etudiants_kholleur;ind_passage ++)
    heure = heure_suivante(heure,duree_kholle,ind_passage,pause,duree_pause_apres_chaque_passage,duree_pause_passages,nb_passages_entre_deux_pauses,nb_etudiants_kholleur);
  //heure.setMinutes(heure.getMinutes()+duree_kholle);
  return heure;
}

function calcule_heure()
{
  let dte = new Date();
  dte.setHours(19);
  dte.setMinutes(14);
  let hh = calcule_heure_debut_preparation(dte,30);
}
 
function calcule_heure_debut_preparation(heure,duree_prepa)
{
  let heure_prepa = new Date(heure);
  heure_prepa.setMinutes(heure.getMinutes() - duree_prepa);
  let hh = heure_prepa.getHours().toString();
  let mn = heure_prepa.getMinutes().toString();
  if (mn.length==1)
      mn = "0"+mn;
  return  hh + "h" + mn;
}

function formate_creneau_horaire(heure,duree_kholle)
{
  let hh_debut = heure.getHours();
  let mn_debut = heure.getMinutes().toString();
  let heure_fin = ajoute_duree_heure(heure,duree_kholle);
  let hh_fin = heure_fin.getHours();
  let mn_fin = heure_fin.getMinutes().toString();
  if (mn_debut.length==1)
      mn_debut = "0"+mn_debut;
  if (mn_fin.length==1)
      mn_fin = "0"+mn_fin;
  return hh_debut + "h" + mn_debut + " - " + hh_fin + "h" + mn_fin;
}

function showSidebarGenerePlanning() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('propose_planning').setTitle("Génération des plannings de khôlles");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
  return htmlOutput ;
}

function showSideBarInfos(titre) {
  const html = HtmlService.createHtmlOutputFromFile('infos').setTitle(titre);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showSidebar_proposition_planning_en_cours() {
 var htmlOutput = HtmlService.createHtmlOutputFromFile('propose_planning').setTitle("Génération des plannings de khôlles");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function proposer_kholleurs_et_eleves_avec_dispo_kholleurs()
{
  
  proposer_kholleurs_et_eleves(true);
}

function proposer_kholleurs_et_eleves_sans_dispo_kholleurs()
{
  
  proposer_kholleurs_et_eleves(false);
}

function determine_repartition_nb_eleves_par_kholleurs(nb_eleves,nb_kholleurs)
{
  let repartition = [];
  for (let i = 0;i < nb_kholleurs;i++)
    repartition.push(0);
  let reste = nb_eleves;
  while (reste != 0)
  {
    for (let i = 0;i < nb_kholleurs && reste != 0;i++)
    {
        repartition[i]++;
        reste--;
    }    
  }
  return repartition;
}

function test_determine_repartition_nb_eleves_par_kholleurs()
{
  for (let nb_eleves = 10; nb_eleves <= 40; nb_eleves ++)
  {
    for (let nb_kholleurs = 1; nb_kholleurs <= nb_eleves; nb_kholleurs++)
    {
       console.log(nb_eleves,nb_kholleurs,determine_repartition_nb_eleves_par_kholleurs2(nb_eleves,nb_kholleurs));
    }
  }
}

function getd()
{
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  const sh_nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  const k = sh_nom_onglet_planning.getRange(3,1,26,5).getValues();
  //console.log("hm");
}

function proposer_kholleurs_et_eleves(avec_dispo_kholleurs)
{
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  const sh_nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  const infos_planning = retourne_infos_from_nom_onglet_planning(nom_onglet_planning);
  const annee = infos_planning["annee"];
  let date_kholle = infos_planning["date_kholle"];
  const date_split = date_kholle.split("/");
  date_kholle = new Date(date_split[2],date_split[1]-1,date_split[0]);
  const discipline = infos_planning["discipline"];
  const nb_examinateurs = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
  let kholleurs_discipline = charge_kholleurs_discipline(annee,date_kholle,discipline,avec_dispo_kholleurs);
  //kholleurs_discipline = shuffle(kholleurs_discipline);
  kholleurs_discipline = retourne_kholleurs_tries(kholleurs_discipline);
  charge_matieres();
  let eleves_annee_discipline = charge_eleves_annee_discipline(annee,discipline);
  let nb_eleves = eleves_annee_discipline.length;
  const repartition_nb_eleves_par_kholleurs = determine_repartition_nb_eleves_par_kholleurs(nb_eleves,nb_examinateurs);
  let historique_eleves = calcule_historique_eleves(annee);
  let histo_kholleurs = calcule_historique_kholleurs();
  //eleves_annee_discipline = shuffle(eleves_annee_discipline);
  let proportion_frequence_meme_kholleur =  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("proportion_frequence_meme_kholleur").getValue();
  let no_ligne=3;
  for (let i = 1;i <= nb_examinateurs && i <= kholleurs_discipline.length; i++)
  {
    let civilite_kholleur = kholleurs_discipline[i-1]["civilite"];
    let prenom_kholleur = kholleurs_discipline[i-1]["prenom"];
    let nom_kholleur = kholleurs_discipline[i-1]["nom"];
    let kholleur = formate_nom_kholleur(civilite_kholleur,prenom_kholleur,nom_kholleur);
    sh_nom_onglet_planning.getRange(no_ligne, 2).setValue(kholleur);
    let eleves_tries_pour_kholleur = retourne_eleves_tries_pour_kholleur(eleves_annee_discipline,historique_eleves,histo_kholleurs,prenom_kholleur,nom_kholleur,proportion_frequence_meme_kholleur);
    for (let ind_eleves = 0; ind_eleves < repartition_nb_eleves_par_kholleurs[i-1]; ind_eleves++)
    {
      let eleve_choisi = eleves_tries_pour_kholleur[0]["eleve"];
      if (sh_nom_onglet_planning.getRange(no_ligne, 3).getValue().substr(0,5) == "Pause")
        no_ligne++;
      sh_nom_onglet_planning.getRange(no_ligne, 3).setValue(eleve_choisi);
      eleves_tries_pour_kholleur.splice(0,1);
      let eleves_annee_discipline_f = eleves_annee_discipline.filter(eleve => eleve !== eleve_choisi);
      eleves_annee_discipline = eleves_annee_discipline_f.slice();
      no_ligne++;
    }
  }
  sh_nom_onglet_planning.autoResizeColumns(1, sh_nom_onglet_planning.getMaxColumns());
  //propage_classement_kholleurs_listes_deroulantes(kholleurs_discipline,nb_examinateurs);
}

function test_proposer_kholleurs_et_eleves_opt()
{
  proposer_kholleurs_et_eleves_opt(false);
}


function proposer_kholleurs_et_eleves_opt(avec_dispo_kholleurs)
{
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  const sh_nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  const nb_examinateurs = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
  const der_ligne = sh_nom_onglet_planning.getRange(LIGNE_MEMO+1,COLONNE_MEMO_NB_ELEVES+nb_examinateurs).getValue();
  const infos_planning = retourne_infos_from_nom_onglet_planning(nom_onglet_planning);
  const annee = infos_planning["annee"];
  let date_kholle = infos_planning["date_kholle"];
  const date_split = date_kholle.split("/");
  date_kholle = new Date(date_split[2],date_split[1]-1,date_split[0]);
  const discipline = infos_planning["discipline"];
  ;
  let kholleurs_discipline = charge_kholleurs_discipline(annee,date_kholle,discipline,avec_dispo_kholleurs);
  //kholleurs_discipline = shuffle(kholleurs_discipline);
  kholleurs_discipline = retourne_kholleurs_tries(kholleurs_discipline);
  charge_matieres();
  let eleves_annee_discipline = charge_eleves_annee_discipline(annee,discipline);
  const nb_eleves = eleves_annee_discipline.length;
  const repartition_nb_eleves_par_kholleurs = determine_repartition_nb_eleves_par_kholleurs(nb_eleves,nb_examinateurs);
  let historique_eleves = calcule_historique_eleves(annee);
  let histo_kholleurs = calcule_historique_kholleurs();
  let donnees_proposees = [];
  //eleves_annee_discipline = shuffle(eleves_annee_discipline);
  let proportion_frequence_meme_kholleur =  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("proportion_frequence_meme_kholleur").getValue();
  let no_ligne=3;
  for (let i = 1;i <= nb_examinateurs && i <= kholleurs_discipline.length; i++)
  {
    let civilite_kholleur = kholleurs_discipline[i-1]["civilite"];
    let prenom_kholleur = kholleurs_discipline[i-1]["prenom"];
    let nom_kholleur = kholleurs_discipline[i-1]["nom"];
    let kholleur = formate_nom_kholleur(civilite_kholleur,prenom_kholleur,nom_kholleur);
    sh_nom_onglet_planning.getRange(no_ligne, 2).setValue(kholleur);
    let eleves_tries_pour_kholleur = retourne_eleves_tries_pour_kholleur(eleves_annee_discipline,historique_eleves,histo_kholleurs,prenom_kholleur,nom_kholleur,proportion_frequence_meme_kholleur);
    for (let ind_eleves = 0; ind_eleves < repartition_nb_eleves_par_kholleurs[i-1]; ind_eleves++)
    {
      let eleve_choisi = eleves_tries_pour_kholleur[0]["eleve"];
      if (sh_nom_onglet_planning.getRange(no_ligne, 3).getValue().substr(0,5) == "Pause")
        no_ligne++;
      sh_nom_onglet_planning.getRange(no_ligne, 3).setValue(eleve_choisi);
      eleves_tries_pour_kholleur.splice(0,1);
      let eleves_annee_discipline_f = eleves_annee_discipline.filter(eleve => eleve !== eleve_choisi);
      eleves_annee_discipline = eleves_annee_discipline_f.slice();
      no_ligne++;
    }
  }
  sh_nom_onglet_planning.autoResizeColumns(1, sh_nom_onglet_planning.getMaxColumns());
  //propage_classement_kholleurs_listes_deroulantes(kholleurs_discipline,nb_examinateurs);
}


function proposer_kholleurs_uniquement()
{
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  const sh_nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  const infos_planning = retourne_infos_from_nom_onglet_planning(nom_onglet_planning)
  const annee = infos_planning["annee"];
  let date_kholle = infos_planning["date_kholle"];
  const date_split = date_kholle.split("/");
  date_kholle = new Date(date_split[2],date_split[1]-1,date_split[0]);
  const discipline = infos_planning["discipline"];
  const nb_examinateurs = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
  let kholleurs_discipline = charge_kholleurs_discipline(annee,date_kholle,discipline,true);
  const nb_eleves = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES).getValue();
  const repartition_nb_eleves_par_kholleurs = determine_repartition_nb_eleves_par_kholleurs(nb_eleves,nb_examinateurs);
  kholleurs_discipline = retourne_kholleurs_tries(kholleurs_discipline);
  let no_ligne=3;
  let no_eleve = 0;
  for (let i = 1;i <= nb_examinateurs; i++)
  {
    let row = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES+i).getValue();
    sh_nom_onglet_planning.getRange(row,2).setValue("");
    for (let ind_eleves = 0; ind_eleves < repartition_nb_eleves_par_kholleurs[i-1]; ind_eleves++)
      {
        if (sh_nom_onglet_planning.getRange(no_ligne, 3).getValue().substr(0,5) == "Pause")
          no_ligne++;
        let cellule_eleve = sh_nom_onglet_planning.getRange(no_ligne, 3);
        cellule_eleve.setValue("");
        no_ligne++;
        no_eleve++;
      }
  }
  for (let i = 1;i <= nb_examinateurs && i <= kholleurs_discipline.length; i++)
  {
    const civilite_kholleur = kholleurs_discipline[i-1]["civilite"];
    let prenom_kholleur = kholleurs_discipline[i-1]["prenom"];
    let nom_kholleur = kholleurs_discipline[i-1]["nom"];
    let kholleur = formate_nom_kholleur(civilite_kholleur,prenom_kholleur,nom_kholleur);
    const no_ligne = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES+i).getValue();
    sh_nom_onglet_planning.getRange(no_ligne, 2).setValue(kholleur);
  }
} 

function lit_donnees_kholleurs_feuille_planning(sheet,nb_examinateurs)
{
  return sheet.getRange(LIGNE_MEMO + 10,COLONNE_MEMO_NB_EXAMINATEURS,nb_examinateurs,2).getValues();
}

function proposer_eleves()
{
  const sh_temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CALCUL_TEMPORAIRES);
  const etat_planning = sh_temp.getRange("C2").getValue();
  if (etat_planning != PLANNING_ECRIT_NON_ENREGISTRE)
  {
     SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue("Il faut au préalable écrire le planning avec les examinateurs choisis");
    showSideBarInfos("Paramètres de la khôlle");
    return;
  }
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  const sh_nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  const infos_planning = retourne_infos_from_nom_onglet_planning(nom_onglet_planning);
  const annee = infos_planning["annee"];
  const discipline = infos_planning["discipline"];
  const date_kholle = infos_planning["date_kholle"];
  const nb_examinateurs = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
  charge_matieres();
  let eleves_annee_discipline = charge_eleves_annee_discipline(annee,discipline);
  const repartition_nb_eleves_par_kholleurs = sh_nom_onglet_planning.getRange(LIGNE_MEMO + 10,COLONNE_MEMO_NB_EXAMINATEURS+1,nb_examinateurs,1).getValues();
  let historique_eleves = calcule_historique_eleves(annee);
  let histo_kholleurs = calcule_historique_kholleurs();
  let proportion_frequence_meme_kholleur =  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("proportion_frequence_meme_kholleur").getValue();
  let no_ligne=3;
  let eleves_deja_choisis = [];
  for (let i = 1;i <= nb_examinateurs; i++)
  {
    const row = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES+i).getValue();
    const civilite_nom_prenom_kholleur = sh_nom_onglet_planning.getRange(row,2).getValue();
    sh_nom_onglet_planning.getRange(row,1).setValue("Salle " + i);
    if (civilite_nom_prenom_kholleur != "")
    {
      let split_civilite_nom_prenom_kholleur = civilite_nom_prenom_kholleur.split(" ");
      let prenom_kholleur = split_civilite_nom_prenom_kholleur[1];
      let nom_kholleur = split_civilite_nom_prenom_kholleur[2];
      let eleves_tries_pour_kholleur = retourne_eleves_tries_pour_kholleur(eleves_annee_discipline,historique_eleves,histo_kholleurs,prenom_kholleur,nom_kholleur,proportion_frequence_meme_kholleur);
      let eleves_tries_pour_kholleur_nom_prenom = [];
      for (let ind_eleve = 0;ind_eleve < eleves_tries_pour_kholleur.length;ind_eleve++)
      {
        let l_eleve = eleves_tries_pour_kholleur[ind_eleve]["eleve"];
        eleves_tries_pour_kholleur_nom_prenom.push(l_eleve);
      }
      for (let ind_eleves = 0; ind_eleves < repartition_nb_eleves_par_kholleurs[i-1]; ind_eleves++)
      {
         if (sh_nom_onglet_planning.getRange(no_ligne, 3).getValue().toString().substr(0,5) == "Pause")
          no_ligne++;
        const passage = sh_nom_onglet_planning.getRange(no_ligne, 5).getValue();
        const eleves_possibles = retourne_eleves_possibles_pour_heure_kholle(eleves_tries_pour_kholleur_nom_prenom,date_kholle,passage,discipline,annee);
        let eleve_choisi = choisit_eleve(passage,eleves_possibles,eleves_deja_choisis);
        eleves_deja_choisis.push(eleve_choisi);
        let cellule_eleve = sh_nom_onglet_planning.getRange(no_ligne, 3);
        cellule_eleve.setValue(eleve_choisi);
        cellule_eleve.setDataValidation(SpreadsheetApp.newDataValidation()
        .setAllowInvalid(true)
        .requireValueInList(eleves_possibles, true)
        .build());
        no_ligne++;
      }
    }
  }
  sh_nom_onglet_planning.autoResizeColumns(2, 8);
  sh_nom_onglet_planning.setColumnWidth(1,150);
}

function preconstruit_eleves_possibles_par_heure(sh_onglet,nb_examinateurs,annee,discipline,date_kholle,eleves_annee_discipline,historique_eleves,histo_kholleurs,proportion_frequence_meme_kholleur,repartition_nb_eleves_par_kholleurs)
{
  let eleves_possibles_par_heure = [];
  const donnees_kholleurs = lit_donnees_kholleurs_feuille_planning(sh_onglet,nb_examinateurs);
  let no_ligne=3;
  for (let i = 1;i <= nb_examinateurs; i++)
  {
    const civilite_nom_prenom_kholleur = donnees_kholleurs[i-1][0];
    const split_civilite_nom_prenom_kholleur = civilite_nom_prenom_kholleur.split(" ");
    const prenom_kholleur = split_civilite_nom_prenom_kholleur[1];
    const nom_kholleur = split_civilite_nom_prenom_kholleur[2];
    let eleves_tries_pour_kholleur = retourne_eleves_tries_pour_kholleur(eleves_annee_discipline,historique_eleves,histo_kholleurs,prenom_kholleur,nom_kholleur,proportion_frequence_meme_kholleur);
    let eleves_tries_pour_kholleur_nom_prenom = [];
      for (let ind_eleve = 0;ind_eleve < eleves_tries_pour_kholleur.length;ind_eleve++)
      {
        let l_eleve = eleves_tries_pour_kholleur[ind_eleve]["eleve"];
        eleves_tries_pour_kholleur_nom_prenom.push(l_eleve);
      }
    for (let ind_eleves = 0; ind_eleves < repartition_nb_eleves_par_kholleurs[i-1]; ind_eleves++)
    {
      if (sh_onglet.getRange(no_ligne, 3).getValue().substr(0,5) == "Pause")
        no_ligne++;
      const passage = sh_onglet.getRange(no_ligne, 5).getValue();
      const eleves_possibles = retourne_eleves_possibles_pour_heure_kholle(eleves_tries_pour_kholleur_nom_prenom,date_kholle,passage,discipline,annee);
      eleves_possibles_par_heure.push({"heure":passage,"kholleur":civilite_nom_prenom_kholleur,"les_eleves":eleves_possibles});
      no_ligne++;
    }
  }
  eleves_possibles_par_heure.sort(
      function(e1, e2) 
      {
        return e1.les_eleves.length - e2.les_eleves.length;
    });
  return eleves_possibles_par_heure;
}

function attribue_heures_eleves(eleves_possibles_par_heure)
{
  let eleves_places = [];
  for (let i = 0; i < eleves_possibles_par_heure.length; i++)
  {

  }
}

function recherche_placement_optimal_eleves(eleves_possibles,pos=0,indice=0,les_eleves_choisis=[])
{
  if (pos != eleves_possibles.length)
  {
    const eleves_possibles_shuffle = eleves_possibles[pos]["les_eleves"];
    const eleve_possible = eleves_possibles_shuffle[indice];
    if (les_eleves_choisis.includes(eleve_possible) )
    {
      if (indice < eleves_possibles[pos]["les_eleves"].length)
        recherche_placement_optimal_eleves(eleves_possibles,pos,indice+1,les_eleves_choisis);
      else
        recherche_placement_optimal_eleves(eleves_possibles,pos+1,0,les_eleves_choisis);
    }
    else
    {
      les_eleves_choisis.push(eleve_possible);
      eleves_possibles[pos]["eleve_choisi"] = eleve_possible;
      recherche_placement_optimal_eleves(eleves_possibles,pos+1,0,les_eleves_choisis);
    }
  }
}

function recherche_placement_optimal_eleves_v2(eleves_possibles,pos=0,les_eleves_choisis=[])
{
  //console.log(pos,les_eleves_choisis.length);
  if (les_eleves_choisis.length == eleves_possibles.length)
    return eleves_possibles;
  const nb_eleves_possibles = eleves_possibles[pos]["les_eleves"].length;
  const les_eleves_choisis_save = [...les_eleves_choisis];
  for (let i = 0; i < nb_eleves_possibles; i++)
  {
    const eleve_possible = eleves_possibles[pos]["les_eleves"][i];
    if (!les_eleves_choisis.includes(eleve_possible) )
    {
      les_eleves_choisis.push(eleve_possible);
      eleves_possibles[pos]["eleve_choisi"] = eleve_possible;
    }
    if (pos < eleves_possibles.length - 1)
      recherche_placement_optimal_eleves_v2(eleves_possibles,pos+1,les_eleves_choisis_save);
    else
      return -1;
  }
}

class Arbre
{
	constructor(racine)
	{
		this.racine = racine;
	}
}

function proposer_eleves_V2()
{
  const sh_temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CALCUL_TEMPORAIRES);
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  const sh_nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  const etat_planning = sh_temp.getRange("C2").getValue();
  if (etat_planning != PLANNING_ECRIT_NON_ENREGISTRE || sh_nom_onglet_planning == null)
  {
     SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue("Il faut au préalable écrire le planning avec les examinateurs choisis");
    showSideBarInfos("Paramètres de la khôlle");
    return;
  }
  const infos_planning = retourne_infos_from_nom_onglet_planning(nom_onglet_planning);
  const annee = infos_planning["annee"];
  const discipline = infos_planning["discipline"];
  let date_kholle = infos_planning["date_kholle"];
  const nb_examinateurs = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
  const donnees_kholleurs = lit_donnees_kholleurs_feuille_planning(sh_nom_onglet_planning,nb_examinateurs);
  charge_matieres();
  const date_kholle_split = date_kholle.split("/");
  date_kholle = new Date(date_kholle_split[2],date_kholle_split[1]-1,date_kholle_split[0]);
  const infos = cherche_si_kholle_meme_semaine_meme_discipline(annee,discipline,date_kholle);
  let eleves_deja_kholles_meme_semaine_meme_discipline;
   if (infos == null)
    eleves_deja_kholles_meme_semaine_meme_discipline = [];
   else
    eleves_deja_kholles_meme_semaine_meme_discipline = retourne_eleves_deja_kholles_meme_semaine_meme_discipline(annee,infos,discipline);
  const eleves_annee_discipline = charge_eleves_annee_discipline(annee,discipline,eleves_deja_kholles_meme_semaine_meme_discipline);
  const repartition_nb_eleves_par_kholleurs = sh_nom_onglet_planning.getRange(LIGNE_MEMO + 10,COLONNE_MEMO_NB_EXAMINATEURS+1,nb_examinateurs,1).getValues();
  let historique_eleves = calcule_historique_eleves(annee);
  let histo_kholleurs = calcule_historique_kholleurs();
  let proportion_frequence_meme_kholleur =  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("proportion_frequence_meme_kholleur").getValue();
  let no_ligne=3;
  const eleves_possibles_par_heure = preconstruit_eleves_possibles_par_heure(sh_nom_onglet_planning,nb_examinateurs,annee,discipline,date_kholle,eleves_annee_discipline,historique_eleves,histo_kholleurs,proportion_frequence_meme_kholleur,repartition_nb_eleves_par_kholleurs);
  const placement_optimal = recherche_placement_optimal_eleves_v3(eleves_possibles_par_heure);
  if (!placement_optimal)
  {
      SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue("Impossible de placer tous les élèves<BR>Changez la date de la khôlle ou bien son heure de début.");
      showSideBarInfos("Placement des étudiants");
    return;    
  }
  for (let i = 1;i <= nb_examinateurs; i++)
  {
      const kholleur = donnees_kholleurs[i-1][0];
      for (let ind_eleves = 0; ind_eleves < repartition_nb_eleves_par_kholleurs[i-1]; ind_eleves++)
      {
         if (sh_nom_onglet_planning.getRange(no_ligne, 3).getValue().substr(0,5) == "Pause")
          no_ligne++;
        const passage = sh_nom_onglet_planning.getRange(no_ligne, 5).getValue();
        const res = cherche_eleve_choisi(passage,eleves_possibles_par_heure,kholleur);
        let cellule_eleve = sh_nom_onglet_planning.getRange(no_ligne, 3);
        if (res.length > 0)
        {
          const eleve_choisi = res[0]["eleve_choisi"];
          cellule_eleve.setValue(eleve_choisi);
          const eleves_possibles = res[0]["les_eleves"];
          if (eleves_possibles != null)
          {
              cellule_eleve.setDataValidation(SpreadsheetApp.newDataValidation()
              .setAllowInvalid(true)
              .requireValueInList(eleves_possibles, true)
              .build());
          }
        }
        else
        {
          msg_log = "Pb pas d'élève choisi pour la kholle : " + annee + " " + discipline + " " + date_kholle + " " + passage + " kholleur : " + kholleur;
          ecrit_log(LOG_WARNING,"proposer_eleves_V2",msg_log);
        }
        no_ligne++;
      }
  }
  sh_nom_onglet_planning.autoResizeColumns(2, 8);
  sh_nom_onglet_planning.setColumnWidth(1,150);
   sh_nom_onglet_planning.getRange('D:D').setHorizontalAlignment('center');
}

function choisit_eleve(passage,eleves_possibles_par_heure,eleves_deja_choisis)
{
  for (const eleves of eleves_possibles_par_heure)
  {
    if (eleves["heure"] == passage)
    {
      for (const eleve of eleves["les_eleves"])
      {
        let deja = false;
        for (const e2 of eleves_deja_choisis)
        {
          if (eleve == e2)
          {
            deja = true;
            break;
          }
        }
        if (!deja)
          return {"eleve":eleve,"eleves_possibles":eleves["les_eleves"]};
      }
    }
  }
  return -1;
}

function cherche_eleve_choisi(passage,eleves_possibles_par_heure,kholleur)
{
  return eleves_possibles_par_heure.filter(row => row["heure"] == passage && row["kholleur"] == kholleur);
}

function test_conflit()
{
  const plage1 = "11h00 - 11h30";
  const plage2 = "11h32 - 12h00";
  conflit_horaire(plage1,plage2);
}

function conflit_horaire(plage1,plage2)
{
  //console.log(plage1,plage2);
  const plage1_split = plage1.split(" - ");
  const plage2_split = plage2.split(" - ");
  const d1 = plage1_split[0];
  const hd1 = new Date();
  hd1.setHours(d1.split("h")[0]);
  hd1.setMinutes(d1.split("h")[1]);
  const f1 = plage1_split[1];
  const fd1 = new Date();
  fd1.setHours(f1.split("h")[0]);
  fd1.setMinutes(f1.split("h")[1]);
  const d2 = plage2_split[0];
  const hd2 = new Date();
  hd2.setHours(d2.split("h")[0]);
  hd2.setMinutes(d2.split("h")[1]);
  const f2 = plage2_split[1];
  const fd2 = new Date();
  fd2.setHours(f2.split("h")[0]);
  fd2.setMinutes(f2.split("h")[1]);
  const diff = (hd2 - fd1);
  if (fd1 < hd2 && (hd2 - fd1) < LIMITE_CONFLIT_DEUX_KHOLLES*60000)
  {
    //console.log(plage1,plage2);
    return true;
  }
  if (fd2 < hd1 && (hd1 - fd2) < LIMITE_CONFLIT_DEUX_KHOLLES*6000)
  {
    //console.log(plage1,plage2);
    return true;
  }
  if (fd1 >= hd2 && fd1 <= fd2)
  {
    //console.log(plage1,plage2);
    return true;
  }
  if (fd2 >= hd1 && fd2 <= fd1)
  {
    //console.log(plage1,plage2);
    return true;
  }
  return false;
}

function retourne_eleves_possibles_pour_heure_kholle(eleves,date_kholle,passage,discipline,classe)
{
  let donnees_histo = sheetToAssociativeArray(FEUILLE_HISTO_ELEVES);
  donnees_histo = donnees_histo.filter(function(row) 
            {
              const date_kholle_formate = formatDate(date_kholle);
              const date_autre_kholle_formate = formatDate(row["Date de la khôlle"]);
              const condition = date_kholle_formate == date_autre_kholle_formate && row["Discipline de la khôlle"] != discipline && row["Classe"] == classe && conflit_horaire(row["Heure de la khôlle"],passage);
              //const condition = date_kholle_formate != date_autre_kholle_formate && row["Discipline de la khôlle"] != discipline && row["Classe"] == classe;
              //if (condition)
              //  console.log(row["Nom de l'élève"],row["Heure de la khôlle"],passage,condition);
                //console.log(date_kholle_formate,date_autre_kholle_formate);
              return condition;
            }); 
  let eleves_possibles = [];
  for (eleve of eleves)
  {
    let deja = false;
    for (donnee of donnees_histo)
    {
      const nom_prenom = donnee["Nom de l'élève"] + " " + donnee["Prénom de l'élève"];
      if (eleve == nom_prenom)
      {
        deja = true;
        break;
      }
    }
    if (!deja)
    {
        eleves_possibles.push(eleve);
    }
  }
  //console.log(eleves_possibles.length);
  return eleves_possibles;
}

function retourne_eleves_tries_pour_kholleur(eleves_annee_discipline,histo_eleves,histo_kholleurs,prenom_kholleur,nom_kholleur,proportion_frequence_meme_kholleur)
{
  let histo_le_kholleur = retourne_histo_kholleur(prenom_kholleur,nom_kholleur,histo_kholleurs);
  let max_passages_eleves;
  let les_eleves_histo = histo_eleves["histo_eleves"];
  let max_score_heure_passage = histo_eleves["max_score_heure_passage"];
  if (histo_le_kholleur == -1)
    max_passages_eleves = 0;
  else
   max_passages_eleves = histo_le_kholleur["max_passages_eleves"];
  let les_eleves_tries = [];
  let nb_passages = 0;
  let l_eleve_trie;
  let nom_prenom_eleve;
  let le_score_pondere_eleve;
  for (let i = 0; i < eleves_annee_discipline.length ; i++)
  {
    let l_eleve = eleves_annee_discipline[i];
    if (histo_le_kholleur == -1)
    {
      nb_passages = 0;
    }
    else
    {
      let deja = false;
      for (let j = 0 ;j < histo_le_kholleur["histo_eleves"].length && !deja; j++)
      {
        nom_prenom_eleve = histo_le_kholleur["histo_eleves"][j]["nom"] + " " + histo_le_kholleur["histo_eleves"][j]["prenom"]; 
        if (nom_prenom_eleve == l_eleve)
        {
          nb_passages =  histo_le_kholleur["histo_eleves"][j]["nb_passages"];
          deja = true;
        }
      }
      if (!deja)
        nb_passages = 0;
    }
    let score_heure_passage = -1;
    for (let j = 0; j < les_eleves_histo.length && score_heure_passage ==-1 ; j++)
    {
      nom_prenom_eleve = les_eleves_histo[j]["nom_eleve"] + " " + les_eleves_histo[j]["prenom_eleve"]
      if (nom_prenom_eleve == l_eleve)
        score_heure_passage = les_eleves_histo[j]["score_heure_passage"];
    }
    if (max_passages_eleves == 0)
    {
      if (max_score_heure_passage==-1)
      {
        le_score_pondere_eleve = 0;
      }
      else
      {
        le_score_pondere_eleve = (max_score_heure_passage - score_heure_passage+1)*50/max_score_heure_passage;
      }
    }
    else
    {
      le_score_pondere_eleve = nb_passages*50/max_passages_eleves*proportion_frequence_meme_kholleur;
      if (max_score_heure_passage>0)
      {
        le_score_pondere_eleve += (max_score_heure_passage - score_heure_passage+1)*50/max_score_heure_passage*(1-proportion_frequence_meme_kholleur);
      }
    }
    l_eleve_trie = {"eleve":l_eleve,"nb_passages":nb_passages,"score_heure_passage":score_heure_passage,"score_pondere":le_score_pondere_eleve};
    les_eleves_tries.push(l_eleve_trie);
  }
  les_eleves_tries.sort(
      function(e1, e2) 
      {
        return e1.score_pondere - e2.score_pondere;
    });
  return les_eleves_tries;
}
  
function retourne_kholleurs_tries(kholleurs)
{
  const proportion_importance_evaluation_kholleur =  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("Proportion_importance_evaluation_kholleur").getValue();
  let max_passages = 0;
  for (const kholleur of kholleurs)
  {
    if (kholleur["nb_passages"] > max_passages)
      max_passages = kholleur["nb_passages"];
  }
   for (const kholleur of kholleurs)
  {
    const le_score_pondere_evaluation = (5-kholleur["evaluation"])*50/5*proportion_importance_evaluation_kholleur;
    let le_score_pondere_nb_passages;
    if (max_passages == 0)
      le_score_pondere_nb_passages = 0;
    else
      le_score_pondere_nb_passages = (max_passages - kholleur["nb_passages"])*50/max_passages*(1-proportion_importance_evaluation_kholleur); 
    const le_score_pondere_kholleur = le_score_pondere_evaluation + le_score_pondere_nb_passages;
    kholleur["score_pondere"] = le_score_pondere_kholleur;
  }

  kholleurs.sort(
      function(k1, k2) 
      {
        return k2.score_pondere - k1.score_pondere;
    });
  return kholleurs;
}

function formate_nom_kholleur(civilite,prenom,nom)
{
  return civilite + " " + nom + " " +  prenom;
}

function propage_classement_kholleurs_listes_deroulantes(kholleurs,nb_examinateurs)
{
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  const sh_nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  let noms_kholleurs = [];
  for (let i = 0;i < kholleurs.length; i ++)
  {
    let nom_kholleur = formate_nom_kholleur(kholleurs[i]["civilite"],kholleurs[i]["prenom"],kholleurs[i]["nom"]);
    noms_kholleurs.push(nom_kholleur)
  }
  for (let ind_examinateur = 1; ind_examinateur <= nb_examinateurs; ind_examinateur++)
  {
     let cellule_kholleur = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES+i).getValue();
     cellule_kholleur.setDataValidation(SpreadsheetApp.newDataValidation()
        .setAllowInvalid(true)
        .requireValueInList(noms_kholleurs, true)
        .build());
  }
}

function classe_kholleurs(annee,discipline)
{
  let kholleurs_classes = charge_historique_kholleurs(annee,discipline);
  kholleurs_classes.sort(
      function(k1, k2) 
      {
        return k1.nb_passages - k2.nb_passages + k1.pos_random - k2.pos_random;
    });
  return kholleurs_classes;

}

function charge_histo()
{
  charge_historique_kholleurs("L2","Economie");
}

function charge_historique_kholleurs(annee,discipline)
{
  let histo_kholleurs = [];
  let sh_histo_kholleurs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_HISTO_KHOLLEURS);
  let nb_kholleurs = 0;
  while (sh_histo_kholleurs.getRange(nb_kholleurs+2,1).getValue() != "")
  {
    let la_discipline = sh_histo_kholleurs.getRange(nb_kholleurs+2,5).getValue();
    let l_annee = sh_histo_kholleurs.getRange(nb_kholleurs+2,1).getValue();
    if (discipline == la_discipline && annee == l_annee)
    {
      let kholleur = 
        {
          "classe":sh_histo_kholleurs.getRange(nb_kholleurs+2,1).getValue(),
          "nom":sh_histo_kholleurs.getRange(nb_kholleurs+2,2).getValue(),
          "prénom":sh_histo_kholleurs.getRange(nb_kholleurs+2,3).getValue(),
          "civilite":sh_histo_kholleurs.getRange(nb_kholleurs+2,4).getValue(),
          "discipline":sh_histo_kholleurs.getRange(nb_kholleurs+2,5).getValue(),
          "nb_passages":sh_histo_kholleurs.getRange(nb_kholleurs+2,6).getValue(),
          "pos_random":Math.random()
        };
    histo_kholleurs.push(kholleur);
    }
    nb_kholleurs++;
  }
  return histo_kholleurs;
}

function retourne_infos_from_nom_onglet_planning(nom_onglet_planning)
{
  const nom_onglet_split = nom_onglet_planning.split(" : ");
  if (nom_onglet_split.length != 2)
  {
    msg_log = "Erreur format nom onglet planning " + nom_onglet_planning;
    ecrit_log(LOG_ERROR,"retourne_infos_from_nom_onglet_planning",msg_log);
    return;
  }
  const infos_split = nom_onglet_split[1].split(" - ");
   if (infos_split.length != 3)
  {
    msg_log = "Erreur format nom onglet planning " + nom_onglet_planning;
    ecrit_log(LOG_ERROR,"retourne_infos_from_nom_onglet_planning",msg_log);
    return;
  }
  const annee = infos_split[0];
  const date_kholle = infos_split[1];
  const discipline = infos_split[2];
  return {"annee":annee,"date_kholle":date_kholle,"discipline":discipline};
}

function seance_deja_enregistree(nom_onglet_planning)
{
  const infos_planning = retourne_infos_from_nom_onglet_planning(nom_onglet_planning);
  const annee = infos_planning["annee"];
  let date_kholle = infos_planning["date_kholle"];
  const date_split = date_kholle.split("/");
  date_kholle = new Date(date_split[2],date_split[1]-1,date_split[0]);
  const discipline = infos_planning["discipline"];
  const sh_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_PLANNING_KHOLLES);
  let donnees_planning = sh_planning.getDataRange().getValues().slice(1);
  donnees_planning = donnees_planning.filter(
    function(row) 
            {
              return row[0] == annee && (row[1] instanceof Date) && row[1] != "" && formatDate(row[1]) == formatDate(date_kholle) && row[2] == discipline && row[3] == "Oui";
            });
    ;
  if (donnees_planning.length != 0)
    return true;
  return false;
}

function enregistre_seance()
{
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  const sh_temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CALCUL_TEMPORAIRES);
  const sh_nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  const etat_planning = sh_temp.getRange("C2").getValue();
  if (etat_planning != PLANNING_ECRIT_NON_ENREGISTRE || sh_nom_onglet_planning == null)
  {
     SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue("Il faut au préalable écrire le planning avec les examinateurs choisis");
    showSideBarInfos("Paramètres de la khôlle");
    return;
  }
  const ctrl = controle_saisies_seance();
  if (ctrl!=true)
  {
     SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue(ctrl);
    showSideBarInfos("Enregistrement du planning de khôlle");
    return;
  }
  if (seance_deja_enregistree(nom_onglet_planning))
  {
    let msg = "Enregistrement impossible : ce planning a déja été enregistré ! <BR><BR>" + nom_onglet_planning;
    SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue(msg);
    showSideBarInfos("Enregistrement du planning de khôlle");
    return;
  }
  msg_sideBar = "Enregistrement en cours .... <BR><BR>" + nom_onglet_planning;
  msg_sideBar += "...<BR><BR>Veuillez patienter ...";
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue(msg_sideBar);
  showSideBarInfos("Enregistrement du planning de khôlle");
  charge_eleves();
  charge_kholleurs();
  const nb_kholleurs = enregistre_seance_histo_kholleurs();
  const nb_eleves = enregistre_seance_histo_eleves();
  
  const infos_planning = retourne_infos_from_nom_onglet_planning(nom_onglet_planning);
  const annee = infos_planning["annee"];
  const date_kholle = infos_planning["date_kholle"];
  const discipline = infos_planning["discipline"];
  const le_professeur_referent = cree_feuille_controle_professeur(annee,date_kholle,discipline);
  let sh_plannings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_PLANNING_KHOLLES);
  let i = 2;
  while (sh_plannings.getRange(i,1).getValue() != "")
  {
    let annee_pl = sh_plannings.getRange(i,1).getValue();
    if (annee == annee_pl)
    {
      let date_kholle_pl = formatDate(sh_plannings.getRange(i,2).getValue());
      if (date_kholle == date_kholle_pl)
      {
        let discipline_pl = sh_plannings.getRange(i,3).getValue();
        if (discipline == discipline_pl)
        {
          sh_plannings.getRange(i,4).setValue("Oui");
          //sh_plannings.activate();
          let msg_termine = "Enregistrement terminé <BR> <BR>";
          if (ENVOI_MAIL_LORS_ENREGISTREMENT)
          {
            msg_termine += "Lien vers la feuille de saisie des notes et des commentaires envoyé par mail aux " + nb_kholleurs + " examinateurs";
            msg_termine += "<BR><BR>Convocations envoyées par mail aux " + nb_eleves + " étudiants";
            msg_termine += "<BR><BR>Lien vers la feuille de contrôle des notes et des commentaires des examinateurs envoyé par mail au professeur référent : " + le_professeur_referent;
          }
          else
          {
            msg_termine += "Les mails n'ont pas été envoyés aux étudiants, ni aux khôlleurs, ni au professeur référent";
          }
          SpreadsheetApp.getActiveSpreadsheet().getRangeByName
          ("message_sidebar").setValue(msg_termine);
          showSideBarInfos("Enregistrement du planning de khôlle");
          return;
        }
      }
    }
    i++;
  }
  console.log("Erreur enregistre_seance : planning non trouvé : " + annee + " - " + date_kholle);
  msg_log = "planning non trouvé : " + annee + " - " + date_kholle;
  ecrit_log(LOG_ERROR,"enregistre_seance",msg_log);
  sh_plannings.activate();
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue("");
  html = showSidebarGenerePlanning();
  sh_temp.getRange("C2").setValue(PLANNING_ENREGISTRE);
}

function controle_saisies_seance()
{
  let ctrl = controle_professeur_referent();
  if (ctrl!=true)
    return ctrl;
  if (!MODE_SIMULATION)
    ctrl = controle_saisies_salles();
  if (ctrl!=true)
    return ctrl;
  ctrl =  controle_saisies_kholleurs();
  if (ctrl!=true)
    return ctrl;
  ctrl =  controle_saisies_eleves();
  if (ctrl!=true)
    return ctrl;
  return true;
}

function controle_professeur_referent()
{
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  const sh_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  if  (sh_planning.getRange(1,8).getValue() == "")
    return "Professeur référent non renseigné";
  return true;
}

function controle_saisies_salles()
{
  let les_salles = new Set();
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  const sh_nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  const nb_examinateurs = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
  for (let ind_examinateur = 1; ind_examinateur <= nb_examinateurs; ind_examinateur++)
  {
    const row = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES+ind_examinateur).getValue();
    const salle = sh_nom_onglet_planning.getRange(row,1).getValue();
    if (salle == "")
      return "Au moins une salle non renseignée";
    else
    {
      if (les_salles.has(salle))
        return "salle " + salle + " renseignée au moins 2 fois";
      else
        les_salles.add(salle);
    }
  }
  return true;
}

function controle_saisies_kholleurs()
{
  let les_kholleurs = new Set();
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  const sh_nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  const nb_examinateurs = sh_nom_onglet_planning.getRange(1,10).getValue();
  for (let ind_examinateur = 1; ind_examinateur <= nb_examinateurs; ind_examinateur++)
  {
    const row = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES+ind_examinateur).getValue();
    const kholleur = sh_nom_onglet_planning.getRange(row,2).getValue();
    if (kholleur == "")
      return "Au moins un examinateur non renseigné";
    else
    {
      if (les_kholleurs.has(kholleur))
        return "examinateur " + kholleur + " renseigné au moins 2 fois";
      else
        les_kholleurs.add(kholleur);
    }
  }
  return true;
}

function controle_saisies_eleves()
{
  let les_eleves = new Set();
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  let sh_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  const nb_eleves = sh_planning.getRange(LIGNE_MEMO+1,COLONNE_MEMO_NB_ELEVES).getValue();
  let no_ligne = 3;
  let no_eleve = 0;
  while (no_eleve < nb_eleves)
  {
    if (sh_planning.getRange(no_ligne,3).getValue().substr(0,5) != 'Pause')
    {
      let nom_prenom_eleve = sh_planning.getRange(no_ligne,3).getValue();
      if (nom_prenom_eleve == "")
        return "au moins un étudiant non renseigné";
      else
      {
        if (les_eleves.has(nom_prenom_eleve))
          return "Etudiant " + nom_prenom_eleve + " renseigné au moins 2 fois";
        else
          les_eleves.add(nom_prenom_eleve);
      }
    }
    no_ligne++;
    no_eleve++;
  } 
  return true;
}

function cree_feuille_notation(no_kholleur,civilite_nom_prenom_kholleur,classe,date_kholle,discipline,liste_eleves_kholle,salle,derniere_heure)
{
  const le_kholleur = civilite_nom_prenom_kholleur["civilité"] + " " + civilite_nom_prenom_kholleur["nom"] + " " + civilite_nom_prenom_kholleur["prénom"];
  const nom_fichier_releve_notes = "Relevé examinateur " + le_kholleur + " - " + discipline + " - " + classe + " - " + date_kholle;
  const annee_scolaire = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("Annee_scolaire").getValue();
  const chemin_feuille_notation = annee_scolaire + "/" + "Notes/" + discipline + "/Examinateurs/" + le_kholleur;
  const rootFolderId = DriveApp.getRootFolder().getId();
  const folder_id = getFolderIdByPath(rootFolderId,chemin_feuille_notation);
  const formule = "=OR(AND(VALUE(ADDRESS(ROW(),COLUMN()))>=0,VALUE(ADDRESS(ROW(),COLUMN()))<=20),UPPER(ADDRESS(ROW(),COLUMN()))=\"ABS\")";
  const rule_note = SpreadsheetApp.newDataValidation().requireFormulaSatisfied(formule).build();
  if (folder_id==null)
    return null;
  const id_feuille_notation = createSpreadsheetInFolder(folder_id, nom_fichier_releve_notes,nom_fichier_releve_notes);
  if (id_feuille_notation != -1)
  {
     const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
     const sh_nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
     sh_nom_onglet_planning.getRange(LIGNE_MEMO+2+no_kholleur,COLONNE_MEMO_NB_EXAMINATEURS).setValue(id_feuille_notation);
     const fichier_releve_notes = SpreadsheetApp.openById(id_feuille_notation);
    fichier_releve_notes.appendRow([nom_fichier_releve_notes]);
    fichier_releve_notes.appendRow(["Etudiant","Heure de passage","Note","Commentaire"]);		
    for (let i = 0; i < liste_eleves_kholle.length; i++)
    {
      const nom_eleve = liste_eleves_kholle[i]["nom"];
      const prenom_eleve = liste_eleves_kholle[i]["prénom"];
      const l_eleve = retourne_eleve(nom_eleve,prenom_eleve);
      const heure_passage = liste_eleves_kholle[i]["heure_passage"];
      if (l_eleve != -1)
      {
        /*
        fichier_releve_notes.getActiveSheet().getRange(i+3,3).setDataValidation(SpreadsheetApp.newDataValidation()
          .setAllowInvalid(false)
          .setHelpText('Saisissez un nombre compris entre 0 et 20')
          .requireNumberBetween(0, 20)
          .build());
          */
        //fichier_releve_notes.getActiveSheet().getRange(i+3,3).setDataValidation(rule_note);
        fichier_releve_notes.appendRow([nom_eleve + " " + prenom_eleve,heure_passage]);	
        if (ENVOI_MAIL_LORS_ENREGISTREMENT)
        {
          const email = l_eleve["mail"];
          if (isValidEmail(email))
            envoie_mail_eleve_convocation_kholle(civilite_nom_prenom_kholleur,liste_eleves_kholle[i],email,nom_fichier_releve_notes,salle);
          else
          {
            msg_log = "Erreur mail élève invalide : " + mail + " eleve : " + nom_eleve + " " + prenom_eleve;
            ecrit_log(LOG_ERROR,"cree_feuille_notation",msg_log);
          }
        }
      }
    }
    fichier_releve_notes.getActiveSheet().getRange(1,1);
    const sheet = fichier_releve_notes.getActiveSheet();
    mef_tableau(sheet,liste_eleves_kholle.length+2,4);
    sheet.getRange('A1:D1').activate().mergeAcross();
    sheet.getActiveRangeList().setHorizontalAlignment('center');
    sheet.setColumnWidth(1, 400);
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
    sheet.getActiveRangeList().setVerticalAlignment('middle');
    sheet.setRowHeights(1, 2 , 60);
    sheet.setRowHeights(3, sheet.getMaxRows() - 2 , 150);
    sheet.setColumnWidth(4, 715);
     sheet.setColumnWidth(3, 60);
    sheet.getRange('D3:D' + sheet.getMaxRows()).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    sheet.getRange('A1:D2').setHorizontalAlignment('center');
    sheet.getRange('B:C').setHorizontalAlignment('center');
    sheet.getRange('A1:D1').setFontSize(14);
    sheet.getRange('A2:D2').setBackground('#ffe599');

    insere_case_a_cocher_validation(fichier_releve_notes);
    if (ENVOI_MAIL_LORS_ENREGISTREMENT)
    {
      envoie_mail_kholleur_feuille_notation(civilite_nom_prenom_kholleur,id_feuille_notation,nom_fichier_releve_notes,salle);
      ecrit_mails_a_envoyer_saisie_notes_kholleur(civilite_nom_prenom_kholleur,classe,date_kholle,discipline,id_feuille_notation,nom_fichier_releve_notes,derniere_heure);
    }
  }
}

function ecrit_mails_a_envoyer_saisie_notes_kholleur(civilite_nom_prenom_kholleur,classe,date_kholle,discipline,id_feuille_notation,nom_fichier_releve_notes,derniere_heure)
{
  const derniere_heure_split = derniere_heure.split("h");
  const hh = parseInt(derniere_heure_split[0]);
  const mn = parseInt(derniere_heure_split[1]);
  let date_derniere_heure = new Date(date_kholle);
  date_derniere_heure.setHours(hh);
  date_derniere_heure.setMinutes(mn + DELAI_ENVOI_FEUILLE_NOTATION_KHOLLEURS_APRES_DERNIERE_KHOLLE);
  const sh_mails_saisie_notes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_MAILS_SAISIE_NOTES_KHOLLEURS);
  sh_mails_saisie_notes.appendRow([classe,discipline,date_kholle,date_derniere_heure,civilite_nom_prenom_kholleur["civilité"],civilite_nom_prenom_kholleur["nom"],civilite_nom_prenom_kholleur["prénom"],id_feuille_notation,nom_fichier_releve_notes]);
}

function insere_case_a_cocher_validation(feuille)
{
  
  const cell = feuille.getRange('E6');
  const validation = cell.getDataValidation();
  if (!(validation && validation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX))
   {
        feuille.getRange('E6').insertCheckboxes();
        feuille.getRange('F6').setValue("Valider la saisie des notes et des commentaires");
  }
  else
  {
    //console.log(cell.getValue());
    return cell.getValue();
  }
}

function supprime_case_a_cocher_validation(feuille)
{
  
  const cell = feuille.getRange('E6');
  const validation = cell.getDataValidation();
  if (validation && validation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX)
   {
      //console.log(feuille.getName());
      feuille.getRange('E6').setDataValidation(null);
      feuille.getRange('E6').setValue("");
      feuille.getRange('F6').setValue("");
  }
}

function cree_feuille_controle_professeur(classe,date_kholle,discipline)
{
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  const sh_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  const dataRange = sh_planning.getDataRange();
  let values = dataRange.getValues().slice(2);
  values = values.filter(function(row) 
              {
                return row[2] != "" && !(row[2].startsWith("Pause"));
              });
  let kholleur = values[0][1];
  for (let i = 0;i < values.length;i++ )
  {
    if (values[i][1] == "")
      values[i][1] = kholleur;
    else
      kholleur = values[i][1];
  }
  const le_professeur_referent = sh_planning.getRange("H1").getValue();
  const annee_scolaire = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("Annee_scolaire").getValue();
  const chemin_feuille_controle_professeur = annee_scolaire + "/" + "Notes/" + discipline + "/Professeurs/" + le_professeur_referent;
  const rootFolderId = DriveApp.getRootFolder().getId();
  const folder_id = getFolderIdByPath(rootFolderId,chemin_feuille_controle_professeur);
  if (folder_id==null)
    return null;
  const date_kholle_split = date_kholle.split("/");
  const date_premiere_kholle = new Date(date_kholle_split[2],date_kholle_split[1]-1,date_kholle_split[0]);
  const autre_date_kholle = cherche_si_kholle_meme_semaine_meme_discipline(classe,discipline,date_premiere_kholle);
  let id_feuille_controle_professeur;
  let nom_fichier_controle_professeur;
  if (autre_date_kholle != null)
  {
    let deux_dates;
    if (date_premiere_kholle < autre_date_kholle)
      deux_dates = formatDate(date_premiere_kholle) + " et " + formatDate(autre_date_kholle);
    else
      deux_dates = formatDate(autre_date_kholle) + " et " + formatDate(date_premiere_kholle);
    nom_fichier_controle_professeur= "Relevé professeur " + le_professeur_referent + " - " + discipline + " - " + classe + " - " + deux_dates;
    const nom_onglet_autre_date_kholle = construit_nom_onglet_planning(classe,formatDate(autre_date_kholle),discipline);
    const sh_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_autre_date_kholle); 
    id_feuille_controle_professeur = sh_planning.getRange(LIGNE_MEMO+2,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
  }
  else
  {
    nom_fichier_controle_professeur = "Relevé professeur " + le_professeur_referent + " - " + discipline + " - " + classe + " - " + date_kholle;
    id_feuille_controle_professeur = createSpreadsheetInFolder(folder_id, nom_fichier_controle_professeur,nom_fichier_controle_professeur); 
  }
  if (validateFileId(id_feuille_controle_professeur))
  {
    const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
    const sh_nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
    sh_nom_onglet_planning.getRange(LIGNE_MEMO+2,COLONNE_MEMO_NB_EXAMINATEURS).setValue(id_feuille_controle_professeur);
    const fichier_controle_professeur = SpreadsheetApp.openById(id_feuille_controle_professeur);
    if (autre_date_kholle == null)
    {
      fichier_controle_professeur.appendRow([nom_fichier_controle_professeur]);
      fichier_controle_professeur.appendRow(["Etudiant","Examinateur","Note","Commentaire"]);
    }
    else
      fichier_controle_professeur.getActiveSheet().getRange(1,1).setValue(nom_fichier_controle_professeur);
    for (const ligne of values)
    {
       fichier_controle_professeur.appendRow([ligne[2],ligne[1],"",""]);
    }		
    insere_case_a_cocher_validation(fichier_controle_professeur);
    const sheet = fichier_controle_professeur.getActiveSheet();
    let donnees = fichier_controle_professeur.getDataRange().getValues().slice(2);
    donnees.sort(
      function(d1, d2) 
      {
        if (d2[2] > d1[2])
          return -1;
        return 1;
    });
    fichier_controle_professeur.getActiveSheet().getRange(3,1,donnees.length,6).setValues(donnees);
    mef_tableau(sheet,donnees.length+2, 4);
    sheet.getRange('A1:D1').activate().mergeAcross();
    sheet.getActiveRangeList().setHorizontalAlignment('center');
    sheet.setColumnWidth(4, 600);
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
    sheet.setRowHeights(3, 8, 50);
    sheet.getActiveRangeList().setVerticalAlignment('middle');

    sheet.getRange('A1:D2').setHorizontalAlignment('center');
    sheet.getRange('C:C').setHorizontalAlignment('center');
    sheet.getRange('A1:D1').setFontSize(14);
    sheet.getRange('A2:D2').setBackground('#ffe599');
  }
  if (ENVOI_MAIL_LORS_ENREGISTREMENT)
    envoie_mail_professeur_controle_notation(le_professeur_referent,id_feuille_controle_professeur,nom_fichier_controle_professeur);
  return le_professeur_referent;
}

function demande_confirmation_disponibilite_kholleur()
{
  const nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("nom_onglet_planning").getValue();
  const sh_temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_CALCUL_TEMPORAIRES);
  const sh_nom_onglet_planning = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom_onglet_planning);
  const etat_planning = sh_temp.getRange("C2").getValue();
  if (etat_planning != PLANNING_ECRIT_NON_ENREGISTRE || sh_nom_onglet_planning == null)
  {
     SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue("Il faut au préalable écrire le planning avec les examinateurs choisis");
    showSideBarInfos("Paramètres de la khôlle");
    return;
  }
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue("Demande de confirmation de disponibilités aux examinateurs ... <br><br>pour la khôlle : " + nom_onglet_planning + "<BR><BR>Veuillez patienter ....");
  showSideBarInfos("Confirmation disponibilité examinateurs");
  const infos_planning = retourne_infos_from_nom_onglet_planning(nom_onglet_planning)
  const annee = infos_planning["annee"];
  const date_kholle = infos_planning["date_kholle"];
  const discipline = infos_planning["discipline"];
  const nb_examinateurs = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_EXAMINATEURS).getValue();
  const donnees_kholleurs = lit_donnees_kholleurs_feuille_planning(sh_nom_onglet_planning,nb_examinateurs);
  for (let i = 1; i <= nb_examinateurs ; i++)
  {
    const startRow = sh_nom_onglet_planning.getRange(LIGNE_MEMO,COLONNE_MEMO_NB_ELEVES+i).getValue();
    const endRow = sh_nom_onglet_planning.getRange(LIGNE_MEMO+1,COLONNE_MEMO_NB_ELEVES+i).getValue();
    const kholleur = donnees_kholleurs[i-1][0];
    const kholleur_split = kholleur.split(" ");
    const civilite_kholleur = kholleur_split[0];
    const nom_kholleur = kholleur_split[1];
    const prenom_kholleur = kholleur_split[2];
    const heure_debut = sh_nom_onglet_planning.getRange(startRow,4).getValue();
    const heure_fin = sh_nom_onglet_planning.getRange(endRow,5).getValue().split(" - ")[1];
    envoie_mail_kholleur_confirmation_disponibilite(civilite_kholleur,prenom_kholleur,nom_kholleur,date_kholle,discipline,annee,heure_debut,heure_fin);
  }
  let msg = "Demande de confirmation terminée";
  msg += "<br><br> " + nb_examinateurs + " mails envoyés aux examinateurs";
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("message_sidebar").setValue(msg);
  showSideBarInfos("Confirmation disponibilité examinateurs");
}


function supprime_tous_les_plannings()
{
  for (const feuille of SpreadsheetApp.getActiveSpreadsheet().getSheets())
  {
    if (feuille.getName().startsWith("Planning Khôlle :"))
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(feuille);
  }
}