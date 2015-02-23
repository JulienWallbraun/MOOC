import java.util.HashMap;


public class Cohorte {

	HashMap<String, Eleve> mapElevesCohorte = new HashMap<String, Eleve>();
	
	public Cohorte(HashMap<String, Eleve> mapElevesCohorte) {
		super();
		this.mapElevesCohorte = mapElevesCohorte;
	}
	
	public void afficherCohorte(){
		if (mapElevesCohorte != null && mapElevesCohorte.size()>0){
			for (Eleve eleve : mapElevesCohorte.values()){
				eleve.afficherDetailEleve();
			}
		}
	}
	
}
