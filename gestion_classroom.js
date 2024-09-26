function init_classroom()
{
  try {
        const nb_cours_supprimes = supprime_cours();
        msg_log = nb_cours_supprimes + " cours de ClassRoom supprimés";
        console.log(msg_log);
        ecrit_log(LOG_INFO,"init_classroom",msg_log);
        charge_matieres();
        msg_log = les_matieres.length +" matières chargées";
        console.log(msg_log);
        ecrit_log(LOG_INFO,"init_classroom",msg_log);
        charge_classes();
        msg_log = les_classes.length + " classes chargées";
        console.log(msg_log);
        ecrit_log(LOG_INFO,"init_classroom",msg_log);
        charge_salles_kholles();
        msg_log = les_salles_kholles.length + " salles chargées";
        console.log(msg_log);
        ecrit_log(LOG_INFO,"init_classroom",msg_log);
        charge_kholleurs();
        msg_log = les_kholleurs.length +" kholleurs chargés";
        console.log(msg_log);
        ecrit_log(LOG_INFO,"init_classroom",msg_log);
        charge_professeurs();
        msg_log = les_professeurs.length + " professeurs chargés";
        console.log(msg_log);
        ecrit_log(LOG_INFO,"init_classroom",msg_log);
        charge_eleves();
        msg_log = les_eleves.length +" élèves chargés";
        console.log(msg_log);
        ecrit_log(LOG_INFO,"init_classroom",msg_log);
        const nb_classrooms = cree_classes_classroom();
        msg_log = nb_classrooms + " classrooms créées";
        console.log(msg_log);
        ecrit_log(LOG_INFO,"init_classroom",msg_log);
        const nb_kholleurs_assignes = assigne_kholleurs_classes_classroom();
        msg_log = nb_kholleurs_assignes + " kholleurs assignés aux classes de classroom";
        console.log(msg_log);
        ecrit_log(LOG_INFO,"init_classroom",msg_log);
        const nb_profs_assignes = assigne_professeurs_classes_classroom();
        msg_log = nb_profs_assignes + " professeurs assignés aux classes de classroom";
        console.log(msg_log);
        ecrit_log(LOG_INFO,"init_classroom",msg_log);
        const nb_eleves_assignes = assigne_eleves_classes_classroom();
        msg_log = nb_eleves_assignes + " élèves assignés aux classes de classroom";
        console.log(msg_log);
        ecrit_log(LOG_INFO,"init_classroom",msg_log);
       } catch (err) {
           console.log('Erreur dans init_classroom :', err.message);
           ecrit_log(LOG_ERROR,"init_classroom",err.message);
       }
}

function cree_classes_classroom()
{
  let nb_classrooms = 0;
  for (let i = 0;i < les_classes.length; i++)
  {
    for (let j = 0; j < les_matieres.length; j ++)
    {
      let annees = les_matieres[j]["annees"];
      if (annees == "L1 et L2" || annees == les_classes[i]["année"])
      {
        var course = {
          name: les_matieres[j]["nom"] +" - "+les_classes[i]["année"],
          section: les_classes[i]["nom"],
          ownerId : 'me'
        };
        try
        {
          Classroom.Courses.create(course);
          nb_classrooms++;
        }
        catch (err)
        {
          msg_log = "Erreur create courses classroom : " + course.name;
          console.log(msg_log);
          ecrit_log(LOG_ERROR,"cree_classes_classroom",msg_log);
        }
      }
    } 
  }
  return nb_classrooms;
}

function assigne_kholleurs_classes_classroom()
{
  const les_cours = Classroom.Courses.list();  
  const cours = les_cours.getCourses();
  let nb_kholleurs_assignes = 0;
  if (cours.length > 0) {
    for (var i = 0; i < cours.length; i++) {
      const un_cours = cours[i];
      const cours_id = un_cours.getId();
      console.log("Course Name:", un_cours.getName());
      const nom_cours = un_cours.getName();
      const split_nom_cours = nom_cours.split(" - ");
      const matiere = split_nom_cours[0];
      const la_matiere = retourne_matiere(matiere);
      const annees = la_matiere["annees"];
      const annee = split_nom_cours[1];
      const ok_annees = (annees == "L1 et L2" || annee == annees);
      for (let j = 0;j < les_kholleurs.length; j++)
      {
        if (matiere == les_kholleurs[j].matiere && ok_annees)
        {
          if (!deja_kholleur_assigne(cours_id))
          {
            assigne_kholleur(cours_id,j);
            nb_kholleurs_assignes++;
          }
        }
      }
    }
  }
  return nb_kholleurs_assignes;
}

function assigne_professeurs_classes_classroom()
{
  const les_cours = Classroom.Courses.list();  
  const cours = les_cours.getCourses();
  let nb_profs_assignes = 0;
  if (cours.length > 0) {
    for (let i = 0; i < cours.length; i++) {
      const un_cours = cours[i];
      const cours_id = un_cours.getId();
      console.log("Course Name:", un_cours.getName());
      const nom_cours = un_cours.getName();
      const split_nom_cours = nom_cours.split(" - ");
      const matiere = split_nom_cours[0];
      const la_matiere = retourne_matiere(matiere);
      const annees = la_matiere["annees"];
      const annee = split_nom_cours[1];
      const ok_annees = (annees == "L1 et L2" || annee == annees);
      for (let j = 0;j < les_professeurs.length; j++)
      {
        if (matiere == les_professeurs[j]["matière"] && ok_annees)
        {
          if (!deja_professeur_assigne(cours_id) && !est_kholleur(les_professeurs[j]["nom"],les_professeurs[j]["prénom"]))
          {
            assigne_professeur(cours_id,j);
            nb_profs_assignes++;
          }           
        }
      }
    }
  }
  return nb_profs_assignes;
}

function est_kholleur(nom,prenom)
{
  for (let i = 0;i < les_kholleurs.length; i++)
  {
    if (les_kholleurs[i]["nom"] == nom && les_kholleurs[i]["prenom"] == prenom)
      return true;
  }
  return false;
}

function assigne_eleves_classes_classroom()
{
  const les_cours = Classroom.Courses.list();  
  const cours = les_cours.getCourses();
  let nb_eleves_assignes = 0;
  if (cours.length > 0) {
    for (let i = 0; i < cours.length; i++) {
      let un_cours = cours[i];
      const cours_id = un_cours.getId();
      console.log("Course Name:", un_cours.getName());
      const nom_cours = un_cours.getName();
      const split_nom_cours = nom_cours.split(" - ");
      const matiere = split_nom_cours[0];
      const classe = split_nom_cours[1];
      let LV = est_LV(matiere);
      for (let j = 0;j < les_eleves.length; j++)
      {
        if (classe == les_eleves[j].classe && (!LV || matiere == les_eleves[j]["lv"]))
        {
          if (!deja_eleve_assigne(cours_id))
          {
            assigne_eleve(cours_id,j);
            nb_eleves_assignes++;
          }
        }
      }
    }
  }
  return nb_eleves_assignes;
}  

function deja_kholleur_assigne(cours_id)
{
  for (let i = 0;i < les_kholleurs[i].length;i++)
  {
    for (let j = 0;j < les_kholleurs[i].cours_assignes.length;j++)
    {
      if (les_kholleurs[i].cours_assignes[j] == cours_id)
        return true;
    }
  }
  return false;
}

function deja_professeur_assigne(cours_id)
{
  for (let i = 0;i < les_professeurs[i].length;i++)
  {
    for (let j = 0;j < les_professeurs[i].cours_assignes.length;j++)
    {
      if (les_professeurs[i].cours_assignes[j] == cours_id)
        return true;
    }
  }
  return false;
}

function deja_eleve_assigne(cours_id)
{
  for (let i = 0;i < les_eleves[i].length;i++)
  {
    for (let j = 0;j < les_eleves[i].cours_assignes.length;j++)
    {
      if (les_eleves[i].cours_assignes[j] == cours_id)
        return true;
    }
  }
  return false;
}

function assigne_kholleur(courseId, ind_kholleur) {
  var resource = {
    userId: les_kholleurs[ind_kholleur].mail
  };
  try
  {
    Classroom.Courses.Teachers.create(resource,courseId);
    Logger.log("Khôlleur " + les_kholleurs[ind_kholleur].mail + " added to course " + getCourseNameById(courseId));
    les_kholleurs[ind_kholleur].cours_assignes.push(courseId); 
  }
  catch (err) 
  {
    console.log('Erreur assigne_kholleur:', err.message);
    let msg_erreur = "Kholleur : " + les_kholleurs[ind_kholleur]["nom"] + " " + les_kholleurs[ind_kholleur]["prenom"];
    console.log(msg_erreur);
    ecrit_log(LOG_ERROR,"assigne_kholleur",err.message + " - " + msg_erreur);
  }
}

function assigne_professeur(courseId, ind_professeur) {
  var resource = {
    userId: les_professeurs[ind_professeur].mail
  };
  try
  {
    Classroom.Courses.Teachers.create(resource,courseId);
    Logger.log("Professeur " + les_professeurs[ind_professeur].mail + " added to course " + getCourseNameById(courseId));
    les_professeurs[ind_professeur].cours_assignes.push(courseId); 
  }
  catch (err) 
  {
    console.log('Erreur assigne_professeur:', err.message);
    let msg_erreur = "Professeur : " + les_professeurs[ind_professeur]["nom"] + " " + les_professeurs[ind_professeur]["prénom"];
    console.log(msg_erreur);
    ecrit_log(LOG_ERROR,"assigne_professeur",err.message + " - " + msg_erreur);
  }
}

function assigne_eleve(courseId, ind_eleve) {
  var resource = {
    userId: les_eleves[ind_eleve].mail
  };
  try
  {
    Classroom.Courses.Students.create(resource,courseId);
    Logger.log("Student " + les_eleves[ind_eleve].mail + " added to course " + getCourseNameById(courseId));
    les_eleves[ind_eleve].cours_assignes.push(courseId);
  }
    catch (err) 
  {
    console.log('Erreur assigne_eleve:', err.message);
    let msg_erreur = "Elève : " + les_eleves[ind_eleve]["nom"] + " " + les_eleves[ind_eleve]["prenom"];
    console.log(msg_erreur);
    ecrit_log(LOG_ERROR,"assigne_eleve",err.message + " - " + msg_erreur);
  }
}

function getCourseNameById(courseId)
{
  let les_cours = Classroom.Courses.list();  
  let cours = les_cours.getCourses();
  if (!cours)
    return;
  for (let i = 0; i < cours.length; i++)
  {
    if (cours[i].getId() == courseId)
      return cours[i].getName() + " - " + cours[i].section;
  }

}

function listCourses() {
      les_cours = []
       const optionalArgs = {
           pageSize: 10, // You can adjust the page size as needed.
           state: "ACTIVE"
       };

       try {
           const response = Classroom.Courses.list();
           const courses = response.courses;
           if (!courses || courses.length === 0) {
               console.log('No courses found.');
               return;
           }
           // Print the course names and IDs of the available courses.
           for (const course of courses) {
               //console.log('%s (%s)', course.name, course.id);
               les_devoirs = listCoursework(course.id)
               les_cours.push({"name":course.name,"id":course.id,"les_devoirs":les_devoirs})
           }
       } catch (err) {
           console.log('Failed with error:', err.message);
       }
       console.log(les_cours)
}

function listCoursework(courseId) {
  // Optional arguments for filtering results (adjust as needed)
  var optionalArgs = {
    pageSize: 10,  // Maximum number of coursework to return
    courseWorkStates: ["PUBLISHED"]  // Filter for published assignments only
  };

  try {
    // Call the API to list coursework for the given course ID
    var coursework = Classroom.Courses.CourseWork.list(courseId);
    var les_devoirs = [];
    // Process the coursework data (replace with your logic)    
    for (var works in coursework.courseWork) {
      for (var work in works)
      //var work = coursework.works[i];
      //sheet.appendRow([work.title, work.creationTime]);
        les_devoirs.push({"title":work.title,"dueDate":work.dueDate})
        console.log("cct",coursework.courseWork[0].title)
    }
  } catch (e) {
    // Handle errors
    Logger.log("Error listing coursework: " + e);
  }
  return les_devoirs;
}

function createAssignment() {
  var articleLink = {
    title: 'SR-71 Blackbird',
    url: 'https://www.lockheedmartin.com/en-us/news/features/history/blackbird.html'
  }
  var materials = [{ link: articleLink }];

  var content = {
    title: 'Supersonic aviation',
    description: 'Read about how the SR-71 Blackbird was built.',
    materials: materials,
    workType: 'ASSIGNMENT',
    state: 'PUBLISHED'
  };

  try {
    const response = Classroom.Courses.list().getCourses();
    const courseId = response[0].getId();
    var courseWork = Classroom.Courses.CourseWork.create(content, courseId);
    Logger.log('Assignment created with ID: ' + courseWork.id);
  } catch (e) {
    Logger.log('Error creating assignment: ' + e);
  }
}

function cree_kholle_classroom(civilite_nom_prenom_kholleur,classe,date_kholle,discipline,liste_eleves_kholle)
{
  const nom_cours = discipline + " - " + classe;
  const cours_id = retourne_id_cours_par_nom(nom_cours);
  if (cours_id != -1)
  {
    const le_kholleur = civilite_nom_prenom_kholleur["civilité"] + " " + civilite_nom_prenom_kholleur["nom"];
    const nom_kholle = le_kholleur + " " + discipline + " " + classe + " " + formatDate(date_kholle);
    try
    {
      const content = 
      {
      title: nom_kholle,
      description: nom_kholle,
      materials: [],
      workType: 'ASSIGNMENT',
      state: 'PUBLISHED'
     };
      var kholle_assignation = Classroom.Courses.CourseWork.create(content, cours_id);
      msg_log = "création kholle Classroom : " + nom_kholle;
      Logger.log(msg_log);
      ecrit_log(LOG_INFO,"cree_kholle_classroom",msg_log);
    }
    catch (err)
    {
      msg_log = "Erreur cree_kholle_classroom : " + nom_kholle + " - " + err;
      console.error(msg_log);
      ecrit_log(LOG_ERROR,"cree_kholle_classroom",msg_log);
    }
    let eleves_ids = [];
    let les_eleves_classroom = Classroom.Courses.Students.list(cours_id).getStudents();
    for (let i = 0; i < liste_eleves_kholle.length; i++)
    {
      const nom_eleve = liste_eleves_kholle[i]["nom"];
      const prenom_eleve = liste_eleves_kholle[i]["prénom"];
      const l_eleve = retourne_eleve(nom_eleve,prenom_eleve);
      if (l_eleve != -1)
      {
        const email = l_eleve["mail"];
        const eleve_classroom = Classroom.Courses.Students.get(cours_id, email);
        const eleve_id = eleve_classroom.getUserId();
        eleves_ids.push(eleve_id);
        //Classroom.Courses.CourseWork.Students.create(cours_id, kholle_assignation.getId(), {
        //userId: eleve_id });
      }
    }
    let les_eleves_a_supprimer = [];
    for (let i = 0;i < les_eleves_classroom.length; i++)
    {
      let trouve = false;
      const id_eleve_classroom = les_eleves_classroom[i].getUserId();
      for (let j = 0;j < eleves_ids.length && !trouve; j++)
      {
        if ( id_eleve_classroom == eleves_ids[j])
          trouve = true;
      }
      if (!trouve)
        les_eleves_a_supprimer.push(id_eleve_classroom);
    }
    const kholle_id =  kholle_assignation.getId();
    try
    {
      const modifyAssigneesOptions = {
        "assigneeMode": "INDIVIDUAL_STUDENTS",
        "modifyIndividualStudentsOptions": {
          "addStudentIds" : eleves_ids,
          "removeStudentIds": les_eleves_a_supprimer 
        }};
      Classroom.Courses.CourseWork.modifyAssignees(
      modifyAssigneesOptions,
        cours_id,   
      kholle_id
      );
    }
    catch (err)
    {
      msg_log = "Erreur cree_kholle_classroom : " + nom_kholle + " - " + err;
      console.error(msg_log);
      ecrit_log(LOG_ERROR,"cree_kholle_classroom",msg_log);
    }
  }
}
  
function retourne_id_cours_par_nom(nom_cours)
{
  const les_cours = Classroom.Courses.list().getCourses();
  for (let i = 0;i < les_cours.length ; i++)
  {
    if (les_cours[i].getName() == nom_cours)
      return les_cours[i].getId();
  }
  msg_log = "Cours Classroom non trouvé : " + nom_cours;
  console.errorl(msg_log);
  ecrit_log(LOG_ERROR,"retourne_id_cours_par_nom",msg_log);
  return -1;
}

function teste_si_proposition_en_cours()
{
  return 
}