import java.io.File;
import java.io.IOException;

import jxl.read.biff.BiffException;
public class Main {
	
	//� changer
	static String path = "C:"+File.separator+"Users"+File.separator+"wallbraun julie"+File.separator+"Documents"+File.separator+"MinesAles"+File.separator+"3M2"+File.separator+"Etude de cas";
	
	public static void main(String[] args) throws BiffException, IOException {
		// TODO Auto-generated method stub

		/* R�cup�ration du classeur Excel (en lecture) */
		File excelFile = new File(path+File.separator+"test.xls");			
		ExtraireInfosXLS EIXLS = new ExtraireInfosXLS(excelFile);

		//ajout des �l�ves incrits en renseignant leur semaine d'inscription
		EIXLS.ajouterEleves(10);
		
		//mise � jour du dernier HW r�ussi par les �l�ves inscrits
		EIXLS.ajouterDernierHWReussiELeves(10);
		
		//affichage
//		EIXLS.afficherMapEleves();
		
		//rangement des �l�ves dans leur cohorte respective
		EIXLS.rangerElevesInscritsDansCohortes();
		
		//affichage
//		EIXLS.afficherDetailCohorte(5,6);
//		EIXLS.afficherTabCohortes();
		
		//cr�ation du fichier TSV contenant les populations de chaque cohorte
		EIXLS.creerCSV(path,"cohortes.tsv");
		
	}

}
