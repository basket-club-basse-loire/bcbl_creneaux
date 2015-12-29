package bcbl.planning.arbitrage;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Planning {

	public static void main(String[] args) {
		String extract = null;
		String output = null;
		String onglet = null;
		String mappingFile = null;
		String licencies = null;
		boolean overwrite = false;

		for (int i = 0; i < args.length; i++) {
			if ("-extract".equals(args[i])) {
				extract = args[++i];
			} else if ("-output".equals(args[i])) {
				output = args[++i];
			} else if ("-onglet".equals(args[i])) {
				onglet = args[++i];
			} else if ("-mapping".equals(args[i])) {
				mappingFile = args[++i];
			} else if ("-licencies".equals(args[i])) {
				licencies = args[++i];
			} else if ("-overwrite".equals(args[i])) {
				overwrite = true;
			}
		}

		if (extract == null) {
			System.err.println("Fichier d'extraction non spécifié");
			System.exit(1);
		}
		if (output == null) {
			System.err.println("Fichier de sortie non spécifié");
			System.exit(1);
		}
		if (onglet == null) {
			System.err.println("Nom de l'onglet non spécifié");
			System.exit(1);
		}

		try {
			Properties mappings = new Properties();
			mappings.load(new FileInputStream(mappingFile));

			HSSFWorkbook extractFbiWb = new HSSFWorkbook(new FileInputStream(
					extract));
			HSSFSheet extractFbiSheet = extractFbiWb.getSheetAt(0);

			// create a new workbook
			Workbook outputWb;

			File fOutput = new File(output);
			if (!overwrite && fOutput.exists()) {
				outputWb = new HSSFWorkbook(new FileInputStream(fOutput));
			} else {
				outputWb = new HSSFWorkbook();
			}
			// create a new file
			FileOutputStream out = new FileOutputStream(output);
			// create a new sheet
			Sheet phaseSheet = outputWb.createSheet();
			outputWb.setSheetName(outputWb.getSheetIndex(phaseSheet), onglet);

			if (licencies != null) {
				HSSFWorkbook licenciesWb = new HSSFWorkbook(
						new FileInputStream(licencies));
				HSSFSheet licenciesSheet = licenciesWb.getSheetAt(0);

				Sheet effectifsSheet = outputWb.createSheet();
				outputWb.setSheetName(outputWb.getSheetIndex(effectifsSheet), "Effectifs");

				new CreationOngletLicencies(licenciesSheet, outputWb,
						effectifsSheet).doIt();
				
			}
			
			new CreationOngletPhase(extractFbiSheet, outputWb, phaseSheet,
					mappings).doIt();

			outputWb.write(out);
			out.close();
			outputWb.close();

			extractFbiWb.close();

		} catch (IOException ioe) {
			ioe.printStackTrace();
		}

	}

}
