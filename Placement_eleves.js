const MAX_SOLUTIONS  = 1;

class Noeud {
    constructor(value, libelle) {
        this.value = value;
        this.libelle = libelle;
        this.children = [];
        this.parent = null;
    }
}

function retourneSolution(noeud) {
    solution.unshift(noeud.value); // Ajoute la valeur du nœud au début de la liste
    if (noeud.parent !== null) {
        retourneSolution(noeud.parent);
    }
}

function deja(noeud, valeur) {
    if (noeud.value === valeur) {
        return true;
    }
    if (noeud.parent === null) {
        return false;
    }
    return deja(noeud.parent, valeur);
}
var solutions = [];

function construireArbre(matrice, niveau, parent) {
	//console.log(niveau,matrice.length);
	if (solutions.length == MAX_SOLUTIONS)
		return;
    if (niveau === matrice.length) {
        retourneSolution(parent);
		solutions.push(solution);
		solution = [];
    } else {
        for (const valeur of matrice[niveau]) {
            const noeud = new Noeud(valeur, niveau + "-" + valeur);
            if (!deja(parent, valeur)) {
                if (parent) {
                    parent.children.push(noeud);
                    noeud.parent = parent;
                }
                construireArbre(matrice, niveau + 1, noeud);
            }
        }
    }
}

var solution = [];
var les_eleves_choisis=[];
function recherche_placement_optimal_eleves_v3(eleves_possibles)
{
  
	eleves_possibles.sort((a, b) => {
	  if (a.kholleur < b.kholleur) return -1;
	  if (a.kholleur > b.kholleur) return 1;
	  return a.heure.localeCompare(b.heure);
	});
	//for (const pl of planning)
	//   console.log(pl);

	let stan = [];
	for (const pl of eleves_possibles)
		stan.push(pl["les_eleves"])

	const racine = new Noeud(0,0);
	construireArbre(stan, 0, racine);
	//console.log("sol",len(solution),len(planning));
	/*
		for (let i=0; i <planning.length ; i++)
			planning[i]["eleve_choisi"] = solution[i+1]; 
  */    
	let no_solution = 0;	
  if (solutions.length == 0)
    return false;
	for (solution of solutions)	
	{
		no_solution++;
		//console.log("--------------------------");
		//console.log("solution n°" + no_solution);
		let i = 1;
		for (sl of eleves_possibles)
		{
			//console.log(sl["kholleur"],sl["heure"],solution[i]);
      sl["eleve_choisi"] = solution[i];
      les_eleves_choisis.push(sl);
			i++;
		}
	}
  return true;
}







function recherche_placement_optimal_eleves_v4(eleves_possibles,pos=0,les_eleves_choisis=[])
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
      recherche_placement_optimal_eleves_v2(eleves_possibles,pos+1,shuffle(les_eleves_choisis_save));
    else
      return -1;
  }
}