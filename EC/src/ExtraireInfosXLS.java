import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ExtraireInfosXLS {
	
	int INDICE_COLONNE_LOGIN = 1;
	int INDICE_LIGNE_PREMIER_ELEVE = 11;
	int INDICE_COLONNE_PREMIER_HW = 4;
	int INDICE_COLONNE_DERNIER_HW = 33;
	
	private ArrayList<Sheet> listeSheets = new ArrayList<Sheet>();
	private Workbook workbook;
	private HashMap<String, Eleve> mapElevesInscrits = new HashMap<String, Eleve>();
	private int[][] tabPopulationCohortes;

	public HashMap<String, Eleve> getListeElevesInscrits() {
		return mapElevesInscrits;
	}

	public void setListeElevesInscrits(HashMap<String, Eleve> listeElevesInscrits) {
		this.mapElevesInscrits = listeElevesInscrits;
	}

	public Workbook getWorkbook() {
		return workbook;
	}

	public void setWorkbook(Workbook workbook) {
		this.workbook = workbook;
	}

	public ArrayList<Sheet> getListeSheets() {
		return listeSheets;
	}

	public void setListeSheets(ArrayList<Sheet> listeSheets) {
		this.listeSheets = listeSheets;
	}

	public ExtraireInfosXLS(File file) throws BiffException, IOException {
		super();
		this.workbook = Workbook.getWorkbook(file);
		for (int i=0; i<workbook.getNumberOfSheets(); i++){
			listeSheets.add(workbook.getSheet(i));
		}		
	}
	
	public boolean noteSuperieureALaMoyenne(String note){
		if (note.equals("1")) return true;
		else if (note.equals("0")) return false;
		else{
			String premiereDecimale = note.substring(2, 3);
			if (premiereDecimale.equals("5") || premiereDecimale.equals("6")|| premiereDecimale.equals("7")|| premiereDecimale.equals("8")|| premiereDecimale.equals("9")) return true;
			else return false;
		}		
	}
	
	public int dernierHWReussiparEleve(int numLigneEleve, Sheet sheet, int dernierHW) throws BiffException, IOException{
		boolean homeworkReussi = false;
		int dernierHWReussi = 0;
		while (dernierHW>0 && !homeworkReussi){
			String note = sheet.getCell(dernierHW+3, numLigneEleve-1).getContents();
			if (noteSuperieureALaMoyenne(note)) {
				homeworkReussi = true;
				dernierHWReussi = dernierHW;
			}
			dernierHW--;
		}
		return dernierHWReussi;
	}
	
	public HashMap<String, Integer> dernierHWReussiPourEnsembleEleves(int numLignePremierEleve, int numLigneDernierEleve, Sheet sheet, int dernierHW) throws BiffException, IOException{
		HashMap<String, Integer> mapDernierHWReussisParLesEleves = new HashMap<String, Integer>();
		for (int i=numLignePremierEleve; i<=numLigneDernierEleve; i++){
			mapDernierHWReussisParLesEleves.put(sheet.getCell(2, i-1).getContents(), dernierHWReussiparEleve(i, sheet, dernierHW));
		}
		return mapDernierHWReussisParLesEleves;
	}
	
	public int semaineInscriptionParEleve(int numLigneEleve, int derniereSemaine) throws BiffException, IOException{
		boolean semaineInscriptionTrouve = false;
		int semaine = 1;
		while (semaine<=derniereSemaine && !semaineInscriptionTrouve){
			Sheet sheet = workbook.getSheet(semaine);
			if (!sheet.getCell(0, numLigneEleve-1).getContents().isEmpty() && sheet.getCell(0, numLigneEleve-1).getContents()!=null){
				semaineInscriptionTrouve = true;
			}
			else {
				semaine++;
			}
		}
		return semaine;
	}
	
	public boolean eleveInscritSemaine(String nomEleve, int numSemaine){
		boolean trouve = false;
		Sheet sheetSemaine = workbook.getSheet(numSemaine);
		int i=0;
		while (i<sheetSemaine.getColumn(INDICE_COLONNE_LOGIN).length && !trouve){
			if (sheetSemaine.getCell(numSemaine, i).getContents().equals(nomEleve)){
				trouve = true;
			}
			i++;
		}
		return trouve;
	}
	
	public int semaineInscriptionEleve(String nomEleve, int derniereSemaineObservee){
		int semaine = 1;
		boolean trouveSemaineInscription = false;
		while (!trouveSemaineInscription && semaine<derniereSemaineObservee){
			if (eleveInscritSemaine(nomEleve, semaine)){
				trouveSemaineInscription = true;
			}
			else{
				semaine++;
			}
		}
		return semaine;
	}

	public void ajouterEleves(int derniereSemaineEtudiee){
		mapElevesInscrits = new HashMap<String, Eleve>();
		//cas semaines normales
		for (int semaine=1; semaine<derniereSemaineEtudiee; semaine++){
			Sheet sheetSemaine = workbook.getSheet(semaine);
			for (int j=INDICE_LIGNE_PREMIER_ELEVE; j<sheetSemaine.getColumn(INDICE_COLONNE_LOGIN ).length; j++){
				String login = sheetSemaine.getCell(INDICE_COLONNE_LOGIN , j).getContents();				
				if (!mapElevesInscrits.containsKey(login)){
//					System.out.println(login+" "+semaine);
					mapElevesInscrits.put(login, new Eleve(login, semaine, -1));
				}				
			}
		}
		//cas particulier derniere semaine
		Sheet sheetDerniereSemaine = workbook.getSheet(derniereSemaineEtudiee);
		for (int j=INDICE_LIGNE_PREMIER_ELEVE; j<sheetDerniereSemaine.getColumn(derniereSemaineEtudiee).length; j++){
			String login = sheetDerniereSemaine.getCell(INDICE_COLONNE_LOGIN+1, j).getContents();			
			if (!mapElevesInscrits.containsKey(login)){
//				System.out.println(login+" "+derniereSemaineEtudiee);
				mapElevesInscrits.put(login, new Eleve(login, derniereSemaineEtudiee, -1));
			}				
		}
	}

	public int dernierHWReussi(Sheet sheet, String login, boolean derniereSemaineEtudiee){
		//cas normal
		int indiceColonneLogin = INDICE_COLONNE_LOGIN;
		int indiceColonnePremierHW = INDICE_COLONNE_PREMIER_HW;
		int indiceColonneDernierHW = INDICE_COLONNE_DERNIER_HW;
		
		//cas derniere semaine
		if (derniereSemaineEtudiee){
			indiceColonneLogin++;
			indiceColonnePremierHW++;
			indiceColonneDernierHW++;
		}
		
		int indiceLigneEleve=INDICE_LIGNE_PREMIER_ELEVE;
		while (!sheet.getCell(indiceColonneLogin, indiceLigneEleve).getContents().equals(login)){
			indiceLigneEleve++;
		}
		int dernierHWReussi = indiceColonneDernierHW;
		while (!noteSuperieureALaMoyenne(sheet.getCell(dernierHWReussi,indiceLigneEleve).getContents()) && dernierHWReussi>=indiceColonnePremierHW){
			dernierHWReussi--;
		}
		dernierHWReussi = dernierHWReussi-indiceColonnePremierHW+1;		
		return dernierHWReussi;
	}
	
	public void ajouterDernierHWReussiELeves(int derniereSemaineEtudiee){
		int semaine = derniereSemaineEtudiee;
		int nbElevesAvecDernierHWReussiTrouve = 0;
		boolean stop = nbElevesAvecDernierHWReussiTrouve<mapElevesInscrits.size();
		
		
		
		//cas derniere semaine
		Sheet sheetDerniereSemaine = getListeSheets().get(derniereSemaineEtudiee);
		for (int i=INDICE_LIGNE_PREMIER_ELEVE; i<sheetDerniereSemaine.getColumn(INDICE_COLONNE_LOGIN+1).length; i++){
			String loginEleve = sheetDerniereSemaine.getCell(INDICE_COLONNE_LOGIN+1, i).getContents();
			if (mapElevesInscrits.get(loginEleve).dernierHWReussi == -1){
				int dernierHWReussiParLEleve = dernierHWReussi(sheetDerniereSemaine, loginEleve, true);
				mapElevesInscrits.put(loginEleve, new Eleve(loginEleve,mapElevesInscrits.get(loginEleve).semaineInscription,dernierHWReussiParLEleve));
				nbElevesAvecDernierHWReussiTrouve++;
//				System.out.println(mapElevesInscrits.get(loginEleve).login+" "+mapElevesInscrits.get(loginEleve).semaineInscription+" "+mapElevesInscrits.get(loginEleve).dernierHWReussi);
			}
		}
		semaine--;
		//cas semaines normales
		while(semaine>=1){
			Sheet sheetSemaine = getListeSheets().get(semaine);
			for (int i=INDICE_LIGNE_PREMIER_ELEVE; i<sheetSemaine.getColumn(INDICE_COLONNE_LOGIN).length; i++){
				String loginEleve = sheetSemaine.getCell(INDICE_COLONNE_LOGIN, i).getContents();
				if (mapElevesInscrits.get(loginEleve).dernierHWReussi == -1){
					int dernierHWReussiParLEleve = dernierHWReussi(sheetSemaine, loginEleve, false);
//					System.out.println("SSSSS "+semaine+" " +loginEleve+ " "+dernierHWReussiParLEleve);
					mapElevesInscrits.put(loginEleve, new Eleve(loginEleve,mapElevesInscrits.get(loginEleve).semaineInscription,dernierHWReussiParLEleve));
					nbElevesAvecDernierHWReussiTrouve++;
				}
			}
			semaine--;
		}		
	}
	
	public void afficherMapEleves(){
		System.out.println("détail de mapElevesInscrits (login, semaine d'inscription, dernier HW réussi");
		if (!mapElevesInscrits.isEmpty()){
			for (Eleve eleve : mapElevesInscrits.values()){
				System.out.println(eleve.login+" "+eleve.semaineInscription+" "+eleve.dernierHWReussi);
			}
		}
	}
	
	public void remplirTabPopulationCohortes(){
		if (!mapElevesInscrits.isEmpty()){
			//on détermine les dimensions du tableau
			int nbMaxSemaine = 0;
			int nbMaxHWReussis = 0;
			for (Eleve eleve : mapElevesInscrits.values()){
				if (eleve.semaineInscription > nbMaxSemaine) nbMaxSemaine = eleve.semaineInscription;
				if (eleve.dernierHWReussi > nbMaxHWReussis) nbMaxHWReussis = eleve.dernierHWReussi;
			}
			tabPopulationCohortes = new int[nbMaxSemaine][nbMaxHWReussis+1];//on peut réussir 0 HW
			//on remplit le tableau
			for (Eleve eleve : mapElevesInscrits.values()){
				tabPopulationCohortes[eleve.semaineInscription-1][eleve.dernierHWReussi]++;
			}
		}
	}
	
	public void afficherTabPopulationCohortes(){
		System.out.println("détail de tabPopulationCohortes");
		for (int i=0; i<tabPopulationCohortes.length; i++){
			for (int j=0; j<tabPopulationCohortes[i].length; j++){
				System.out.println("tabPopulationCohortes(semaine d'inscription = "+(i+1)+", dernier HW réussi = "+j+") = "+tabPopulationCohortes[i][j]);
			}
		}
	}
}
