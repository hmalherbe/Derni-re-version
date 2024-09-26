const AVEC_ANNUAIRE_PARTAGE = true;
const ID_ANNUAIRE = "1q-qU7908nuzvFB1NbkAiqvBnlnejLFyQPcA3GwGGpeY"
const FEUILLE_LISTE_KHOLLEURS = "Khôlleurs";
const PLAGE_SALLES_KHOLLES = "Salles_kholles";
const FEUILLE_HISTO_KHOLLEURS = "Historique kholleurs";
const FEUILLE_LISTE_PROFESSEURS = "Professeurs";
const FEUILLE_LISTE_ELEVES = "Etudiants L2";
const FEUILLE_PARAMETRES = "Params";
const FEUILLE_CALCUL_TEMPORAIRES = "Temp";
const FEUILLE_HISTO_ELEVES = "Historique élèves";
const FEUILLE_CHOIX_KHOLLEURS = "Choix examinateurs";
//const FEUILLE_PLANNING_KHOLLES = "Planning test";
const FEUILLE_PLANNING_KHOLLES = "Planning khôlles";
const FEUILLE_LOGS = "Logs";
const FEUILLE_CONFIRMATION_DISPO_KHOLLEURS = "Forms dispo kholleurs";
const FEUILLE_TCD_HISTO_ELEVES = "TCD Histo Elèves";
const FEUILLE_TCD_HISTO_KHOLLEURS_ELEVES = "TCD Histo Khôlleurs - Elèves";
const FEUILLE_MAILS_SAISIE_NOTES_KHOLLEURS = "Mails saisie notes khôlleurs";
const FEUILLE_SPOOLER_ENVOIS_RESULTATS_ELEVES = "Spooler résultats étudiants";
const JOURS_SEMAINE =["lundi","mardi","mercredi","jeudi","vendredi"];
const WEEK_DAY = ["dimanche","lundi","mardi","mercredi","jeudi","vendredi","samedi"];
const LOG_ERROR = 0;
const LOG_WARNING = 1;
const LOG_INFO = 2;
const LIBELLES_LOGS = ["Erreur","Avertissement","Infos"];
const NIVEAU_LOG = LOG_INFO;
const AVEC_CLASSROOM = false;
const DOSSIER_MODELES = "Modèles";
const NOM_FICHIER_MODELE_RELEVE_NOTES_EXAMINATEURS ="Modèle relevé notes examinateurs";
const COULEUR_NOTES_ET_COMMENTAIRES_KHOLLEURS_VALIDEES = '#13bc3c';
const COULEUR_NOTES_ET_COMMENTAIRES_KHOLLEURS_NON_VALIDEES = '#33ddff';
const COULEUR_CONFIRMATION_DISPONIBILITE_KHOLLEUR_OUI = '#13bc3c';
const COULEUR_CONFIRMATION_DISPONIBILITE_KHOLLEUR_NON = '#c91212';
const LIGNE_MEMO = 1
const COLONNE_MEMO_NB_EXAMINATEURS = 100;
const COLONNE_MEMO_NB_ELEVES = 101;
const ENVOI_NOTE_ET_COMMENTAIRE = "Note et commentaire";
const ENVOI_NOTE_COMMENTAIRE_UNIQUEMENT = "Commentaire uniquement";
const FEUILLE_CALENDRIER = "Calendrier";
const FEUILLE_PLANNING_TEST = "Planning test";

const ETAT_KHOLLEURS_NON_PROPOSES = 0;
const ETAT_KHOLLEURS_PROPOSES = 1;
const PLANNING_ECRIT_NON_ENREGISTRE = 2;
const PLANNING_ENREGISTRE = 3;
const DELAI_ENVOI_FEUILLE_NOTATION_KHOLLEURS_APRES_DERNIERE_KHOLLE = 120;
const LIMITE_CONFLIT_DEUX_KHOLLES = 40;

const ENVOI_MAIL_LORS_ENREGISTREMENT = false;

const EMAILS_REELS = false;

const MAX_KHOLLEURS_PAR_SEANCE = 5;

const MODE_SIMULATION = true;

var colonne_mail_prof;
var colonne_mail_kholleur; 
var colonne_mail_eleve;
var msg_sideBar;
if (EMAILS_REELS)
{
  colonne_mail_prof = 9;
  colonne_mail_kholleur = 12; 
  colonne_mail_eleve = 6;
}
else
{
  colonne_mail_prof = 6;
  colonne_mail_kholleur = 5; 
  colonne_mail_eleve = 3;
}


