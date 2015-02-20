import java.io.File;
import java.io.IOException;

import jxl.Workbook;
import jxl.read.biff.BiffException;
public class Main {

	public static void main(String[] args) throws BiffException, IOException {
		// TODO Auto-generated method stub


		Workbook workbook = null;
		try {
			/* Récupération du classeur Excel (en lecture) */
			File excelFile = new File("C:"+File.separator+"Users"+File.separator+"wallbraun julie"+File.separator+"Documents"+File.separator+"MinesAles"+File.separator+"3M2"+File.separator+"Etude de cas"+File.separator+"test.xls");			
			ExtraireInfosXLS EIXLS = new ExtraireInfosXLS(excelFile);
			
			EIXLS.ajouterEleves(10);
			EIXLS.ajouterDernierHWReussiELeves(10);
//			System.out.println(EIXLS.dernierHWReussi(EIXLS.getWorkbook().getSheet(10), "Peahi", true));
			
			EIXLS.afficherMapEleves();
			EIXLS.remplirTabPopulationCohortes();
			EIXLS.afficherTabPopulationCohortes();
			
			
//			for (int i=5000; i>=4500; i--){
//				System.out.println("semaine inscription numéro "+i+" : "+EIXLS.semaineInscriptionEleve(EIXLS.getWorkbook().getSheet(10).getCell(2, i).getContents(),10));
//			}
		} 
		catch (BiffException e) {
			e.printStackTrace();
		} 
		catch (IOException e) {
			e.printStackTrace();
		} 
		finally {
			if(workbook!=null){
				/* On ferme le worbook pour libérer la mémoire */
				workbook.close(); 
			}
		}
	}

}
