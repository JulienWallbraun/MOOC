import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.HashMap;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import com.google.gson.Gson;

public class ExtraireInfosXLS {
	
	//paramètres du classeur excel à renseigner
	int INDICE_COLONNE_LOGIN = 1;
	int INDICE_COLONNE_EMAIL = 0;
	int INDICE_LIGNE_PREMIER_ELEVE = 11;
	int INDICE_COLONNE_PREMIER_HW = 4;
	int INDICE_COLONNE_DERNIER_HW = 33;
	
	Workbook workbook;
	HashMap<String, Eleve> mapElevesInscrits = new HashMap<String, Eleve>();
	Cohorte[][] tabCohortes;

	public ExtraireInfosXLS(File file) throws BiffException, IOException {
		super();
		this.workbook = Workbook.getWorkbook(file);	
	}
	
	/**détermine si une note donnée en paramètre sous la forme d'une string est supérieure ou égale à la moyenne
	 * 
	 * @param note
	 * @return true si la note est supérieure ou égale à 0,5, false sinon
	 */
	public boolean noteSuperieureALaMoyenne(String note){
		if (note.equals("1")) return true;
		else if (note.equals("0")) return false;
		else{
			String premiereDecimale = note.substring(2, 3);
			if (premiereDecimale.equals("5") || premiereDecimale.equals("6")|| premiereDecimale.equals("7")|| premiereDecimale.equals("8")|| premiereDecimale.equals("9")) return true;
			else return false;
		}		
	}

	/**
	 * ajoute les élèves s'étant inscrits au MOOC dans mapElvesInscrits avec leur date d'inscrition en parcourant les semaines une à une jusqu'à la dernière semaine
	 * le derb=nier devoir réussi par les élèves est initialisé à -1
	 * @param derniereSemaineEtudiee
	 */
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

	/**
	 * retourne, pour un élève et une semaine donnée, son dernier devoir réussi au moment de la semaine en cours
	 * @param sheet
	 * @param login
	 * @param derniereSemaineEtudiee
	 * @return
	 */
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
	
	/**
	 * renseigne dans mapElevesInscrits le dernier devoir réussi par les élèves présents dans la liste des élèves inscrits
	 * @param derniereSemaineEtudiee
	 */
	public void ajouterDernierHWReussiELeves(int derniereSemaineEtudiee){
		int semaine = derniereSemaineEtudiee;
		int nbElevesAvecDernierHWReussiTrouve = 0;
		
		//cas derniere semaine
		Sheet sheetDerniereSemaine = workbook.getSheet(derniereSemaineEtudiee);
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
			Sheet sheetSemaine = workbook.getSheet(semaine);
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
	
	/**
	 * retourne la cohorte correspondant à la semaine d'inscription et au dernier TD donné en paramètre
	 * @param numSemaine
	 * @param dernieHWReussi
	 * @return
	 */
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
	
	/**
	 * range les élèves contenus dans mapElevesInscrits dans la cohorte correspondante dans le tableau tabCohortes
	 */
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
	
	public void afficherDetailCohorte(int semaineInscription, int dernierHWReussi){
		tabCohortes[semaineInscription-1][dernierHWReussi].afficherCohorte();
	}
	
	/**
	 * créée un fichier CSV contenant les effectifs de chaque cohorte
	 * @param path
	 * @param name
	 * @throws IOException
	 */
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
		//création ligne cohortes n'ayant réussi aucun HW
		fileWriter.append(newLine+"aucun");
		for (int i=1; i<=nbSemaines; i++){
			fileWriter.append("\t"+Integer.toString(tabCohortes[i-1][0].mapElevesCohorte.size()));
		}
		//création ligne somme cohortes ayant réussi au moins 1 HW
		fileWriter.append(newLine+"autres");
		for (int j=1; j<=nbSemaines; j++){
			int nbEleves = 0;
			for (int i=1; i< nbMaxHWReussis; i++){
				nbEleves = nbEleves + tabCohortes[j-1][i].mapElevesCohorte.size();
			}
			fileWriter.append("\t"+Integer.toString(nbEleves));
		}
		//création autres lignes (données)		
		for (int i=1; i< nbMaxHWReussis; i++){
			fileWriter.append(newLine+"TD"+Integer.toString(i));
			for (int j=1; j<=nbSemaines; j++){
				fileWriter.append("\t"+Integer.toString(tabCohortes[j-1][i].mapElevesCohorte.size()));
			}
		}
		fileWriter.close();
	}
	
	/**
	 * créée un fichier JSON repertoriant tous les élèves s'étant inscrits au MOOC, avec leur login, leur email, leur semaine d'inscription et leur dernier devoir réussi
	 * @param path
	 * @param name
	 * @throws IOException
	 */
	public void creerJSON(String path, String name) throws IOException{
		FileWriter fileWriter = new FileWriter(path+File.separator+name);
		fileWriter.append(new Gson().toJson(mapElevesInscrits));
		fileWriter.close();
	}
}
