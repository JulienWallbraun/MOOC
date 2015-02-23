import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ExtraireInfosXLS {
	
	int INDICE_COLONNE_LOGIN = 1;
	int INDICE_COLONNE_EMAIL = 0;
	int INDICE_LIGNE_PREMIER_ELEVE = 11;
	int INDICE_COLONNE_PREMIER_HW = 4;
	int INDICE_COLONNE_DERNIER_HW = 33;
	
	ArrayList<Sheet> listeSheets = new ArrayList<Sheet>();
	Workbook workbook;
	HashMap<String, Eleve> mapElevesInscrits = new HashMap<String, Eleve>();
	Cohorte[][] tabCohortes;

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
				String email = sheetSemaine.getCell(INDICE_COLONNE_EMAIL, j).getContents();	
				if (!mapElevesInscrits.containsKey(login)){
					mapElevesInscrits.put(login, new Eleve(login, email, semaine, -1));
				}				
			}
		}
		//cas particulier derniere semaine
		Sheet sheetDerniereSemaine = workbook.getSheet(derniereSemaineEtudiee);
		for (int j=INDICE_LIGNE_PREMIER_ELEVE; j<sheetDerniereSemaine.getColumn(derniereSemaineEtudiee).length; j++){
			String login = sheetDerniereSemaine.getCell(INDICE_COLONNE_LOGIN+1, j).getContents();
			String email = sheetDerniereSemaine.getCell(INDICE_COLONNE_EMAIL+1, j).getContents();	
			if (!mapElevesInscrits.containsKey(login)){
				mapElevesInscrits.put(login, new Eleve(login, email, derniereSemaineEtudiee, -1));
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
		
		//cas derniere semaine
		Sheet sheetDerniereSemaine = listeSheets.get(derniereSemaineEtudiee);
		for (int i=INDICE_LIGNE_PREMIER_ELEVE; i<sheetDerniereSemaine.getColumn(INDICE_COLONNE_LOGIN+1).length; i++){
			String loginEleve = sheetDerniereSemaine.getCell(INDICE_COLONNE_LOGIN+1, i).getContents();
			String emailEleve = sheetDerniereSemaine.getCell(INDICE_COLONNE_EMAIL+1, i).getContents();
			if (mapElevesInscrits.get(loginEleve).dernierHWReussi == -1){
				int dernierHWReussiParLEleve = dernierHWReussi(sheetDerniereSemaine, loginEleve, true);
				mapElevesInscrits.put(loginEleve, new Eleve(loginEleve,emailEleve,mapElevesInscrits.get(loginEleve).semaineInscription,dernierHWReussiParLEleve));
				nbElevesAvecDernierHWReussiTrouve++;
//				System.out.println(mapElevesInscrits.get(loginEleve).login+" "+mapElevesInscrits.get(loginEleve).semaineInscription+" "+mapElevesInscrits.get(loginEleve).dernierHWReussi);
			}
		}
		semaine--;
		//cas semaines normales
		while(semaine>=1 && nbElevesAvecDernierHWReussiTrouve<=mapElevesInscrits.size()){
			Sheet sheetSemaine = listeSheets.get(semaine);
			for (int i=INDICE_LIGNE_PREMIER_ELEVE; i<sheetSemaine.getColumn(INDICE_COLONNE_LOGIN).length; i++){
				String loginEleve = sheetSemaine.getCell(INDICE_COLONNE_LOGIN, i).getContents();
				String emailEleve = sheetSemaine.getCell(INDICE_COLONNE_EMAIL, i).getContents();
				if (mapElevesInscrits.get(loginEleve).dernierHWReussi == -1){
					int dernierHWReussiParLEleve = dernierHWReussi(sheetSemaine, loginEleve, false);
//					System.out.println("SSSSS "+semaine+" " +loginEleve+ " "+dernierHWReussiParLEleve);
					mapElevesInscrits.put(loginEleve, new Eleve(loginEleve,emailEleve,mapElevesInscrits.get(loginEleve).semaineInscription,dernierHWReussiParLEleve));
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
				System.out.println(eleve.login+" "+eleve.email+" "+eleve.semaineInscription+" "+eleve.dernierHWReussi);
			}
		}
	}
	
	public HashMap<String, Eleve> mapElevesInscritsSemaine(int numSemaine){
		HashMap<String, Eleve> mapElevesInscritsSemaine = new HashMap<String, Eleve>();
		for (Eleve eleve : mapElevesInscrits.values()){
			if (eleve.semaineInscription == numSemaine){
				mapElevesInscritsSemaine.put(eleve.login, eleve);
			}
		}
		return mapElevesInscritsSemaine;
	}
	
	public HashMap<String, Eleve> mapElevesDernierHWReussi(int dernieHWReussi){
		HashMap<String, Eleve> mapElevesDernierHWReussi = new HashMap<String, Eleve>();
		for (Eleve eleve : mapElevesInscrits.values()){
			if (eleve.dernierHWReussi == dernieHWReussi){
				mapElevesDernierHWReussi.put(eleve.login, eleve);
			}
		}
		return mapElevesDernierHWReussi;
	}
	
	public Cohorte cohorte(int numSemaine, int dernieHWReussi){
		Cohorte cohorte = new Cohorte(new HashMap<String, Eleve>());
		for (Eleve eleve : mapElevesInscrits.values()){
			if (eleve.semaineInscription == numSemaine && eleve.dernierHWReussi == dernieHWReussi){
				if (cohorte==null){
					cohorte = new Cohorte(new HashMap<String, Eleve>());
				}
				cohorte.mapElevesCohorte.put(eleve.login, eleve);
			}
		}
		return cohorte;
	}
	
	public void rangerElevesInscritsDansCohortes(){
		if (!mapElevesInscrits.isEmpty()){
			//on détermine les dimensions du tableau
			int nbMaxSemaine = 0;
			int nbMaxHWReussis = 0;
			for (Eleve eleve : mapElevesInscrits.values()){
				if (eleve.semaineInscription > nbMaxSemaine) nbMaxSemaine = eleve.semaineInscription;
				if (eleve.dernierHWReussi > nbMaxHWReussis) nbMaxHWReussis = eleve.dernierHWReussi;
			}
			tabCohortes = new Cohorte[nbMaxSemaine][nbMaxHWReussis+1];//on peut réussir 0 HW
			//on remplit les cohortes du tableau
			for (int semaine=0; semaine<tabCohortes.length; semaine++){
				for (int homework=0; homework<tabCohortes[semaine].length; homework++){
					tabCohortes[semaine][homework] = cohorte(semaine+1, homework);
				}
			}
		}
	}
	
	public void afficherTabCohortes(){
		System.out.println("détail tableau cohortes");
		for (int i=0; i<tabCohortes.length; i++){
			for (int j=0; j<tabCohortes[i].length; j++){
				if (tabCohortes[i][j] != null){
					for (Eleve eleve : tabCohortes[i][j].mapElevesCohorte.values()){
						System.out.println("cohorte(semaine="+(i+1)+",dernierHWReussi="+j+") : "+eleve.login+" "+eleve.semaineInscription+" "+eleve.dernierHWReussi);
					}
				}
			}
		}
	}
	
	public void creerCSV(String path, String name) throws IOException{
		int nbSemaines = tabCohortes.length;
		int nbMaxHWReussis = tabCohortes[0].length;
		String newLine = System.getProperty("line.separator");
		FileWriter fileWriter = new FileWriter(path+File.separator+name);
		//création 1ère ligne (titres colonnes)
		fileWriter.append("TDs");
		for (int i=1; i<=nbSemaines; i++){
			fileWriter.append("\t"+"Semaine"+Integer.toString(i));
		}		
		//création autres lignes (données)
		for (int i=0; i< nbMaxHWReussis; i++){
			fileWriter.append(newLine+"TD"+Integer.toString(i));
			for (int j=1; j<=nbSemaines; j++){
				fileWriter.append("\t"+Integer.toString(tabCohortes[j-1][i].mapElevesCohorte.size()));
			}
		}
		fileWriter.close();
	}
}
