import java.io.File;
import java.io.IOException;
import java.util.HashMap;

import jxl.Workbook;
import jxl.read.biff.BiffException;
public class Main {

	public static void main(String[] args) throws BiffException, IOException {
		// TODO Auto-generated method stub


		/* Récupération du classeur Excel (en lecture) */
		File excelFile = new File("C:"+File.separator+"Users"+File.separator+"wallbraun julie"+File.separator+"Documents"+File.separator+"MinesAles"+File.separator+"3M2"+File.separator+"Etude de cas"+File.separator+"test.xls");			
		ExtraireInfosXLS EIXLS = new ExtraireInfosXLS(excelFile);

		EIXLS.ajouterEleves(10);
		EIXLS.ajouterDernierHWReussiELeves(10);			
		//EIXLS.afficherMapEleves();
		EIXLS.rangerElevesInscritsDansCohortes();
		EIXLS.afficherTabCohortes();

		
	}

}
