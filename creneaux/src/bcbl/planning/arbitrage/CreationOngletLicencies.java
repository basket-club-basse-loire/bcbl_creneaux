package bcbl.planning.arbitrage;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class CreationOngletLicencies {

	public final static String BCBL = "BASKET CLUB BASSE LOIRE";
	public final static String ENTENTE = "LOIRE ACHENEAU BASKET";
	public final static String EXEMPT = "Exempt";

	public final static int DIVISION_COLIDX = 0;
	public final static int EQUIPE1_COLIDX = 2;
	public final static int EQUIPE2_COLIDX = 3;
	public final static int DATE_RENCONTRE_COLIDX = 4;
	public final static int HEURE_RENCONTRE_COLIDX = 5;

	private HSSFSheet licenciesSheet;
	private Workbook outputWb;
	private Sheet effectifsSheet;

	int rowLine = 0;

	private CellStyle headerCS;
	private CellStyle contentCS;
	private CellStyle tableRowCS;
	private CellStyle tableBottomCS;
	private CellStyle tableHeaderCS;
	private CellStyle col0CS;
	private Font femininFont;
	private Font masculinFont;
	private Font exterieurFont;

	public static class Licencie {
		public String nomPrenom;
		public String email;
	}

	public class ComparatorLicencie implements Comparator<Licencie> {
		@Override
		public int compare(Licencie o1, Licencie o2) {
			return o1.nomPrenom.compareTo(o2.nomPrenom);
		}
	}

	public CreationOngletLicencies(HSSFSheet licenciesSheet, Workbook outputWb,
			Sheet outputSheet) {
		this.licenciesSheet = licenciesSheet;
		this.outputWb = outputWb;
		this.effectifsSheet = outputSheet;
	}

	public void doIt() throws IOException {
		int rows = licenciesSheet.getPhysicalNumberOfRows();

		Calendar currentWeek = null;
		List<HSSFRow> rowsParWeekend = new ArrayList<HSSFRow>();

		HashMap<String, List<Licencie>> seniors = new HashMap<String, List<Licencie>>();
		HashMap<String, List<Licencie>> jeunes = new HashMap<String, List<Licencie>>();
		HashMap<String, List<Licencie>> autres = new HashMap<String, List<Licencie>>();

		// La 1ere ligne contient les headers
		for (int r = 1; r < rows; r++) {
			HSSFRow row = licenciesSheet.getRow(r);
			if (row != null) {
				Cell cellCategorie = row.getCell(8);
				if (cellCategorie == null) {
					continue;
				}
				String categorie = cellCategorie.getStringCellValue();

				Cell cellEquipe = row.getCell(9);
				if (cellEquipe == null) {
					continue;
				}
				String equipe = cellEquipe.getStringCellValue();

				Licencie licencie = new Licencie();
				licencie.nomPrenom = row.getCell(5).getStringCellValue() + " "
						+ row.getCell(6).getStringCellValue();
				licencie.email = row.getCell(16).getStringCellValue();

				HashMap<String, List<Licencie>> map;
				if (categorie.startsWith("S")) {
					map = seniors;
				} else if (categorie.startsWith("U")) {
					map = jeunes;
				} else {
					map = autres;
				}

				String cle = categorie + " " + equipe;

				cle = cle.trim();
				if (cle.isEmpty()) {
					cle = "Autres";
				}

				List<Licencie> membres = map.get(cle);
				if (membres == null) {
					membres = new ArrayList<Licencie>();
					map.put(cle, membres);
				}
				membres.add(licencie);
			}
		}

		Comparator<String> decroissant = new Comparator<String>() {
			public int compare(String o1, String o2) {
				return o1.compareTo(o2) * -1;
			}
		};

		Comparator<String> croissant = new Comparator<String>() {
			public int compare(String o1, String o2) {
				return o1.compareTo(o2);
			}
		};

		rowLine = 7;
		Row outputRow = effectifsSheet.createRow(rowLine++);
		Cell cell = outputRow.createCell(2);
		cell.setCellValue("Phase 1");
		cell = outputRow.createCell(3);
		cell.setCellValue("Phase 2");
		cell = outputRow.createCell(4);
		cell.setCellValue("Total");

		outputRow = effectifsSheet.createRow(rowLine++);
		cell = outputRow.createCell(1);
		cell.setCellValue("Seniors");
		Row seniorsRecap = outputRow;
		outputRow = effectifsSheet.createRow(rowLine++);
		cell = outputRow.createCell(1);
		cell.setCellValue("Jeunes");
		Row jeunesRecap = outputRow;
		outputRow = effectifsSheet.createRow(rowLine++);
		cell = outputRow.createCell(1);
		cell.setCellValue("Autres");
		Row autresRecap = outputRow;

		rowLine++;

		outputRow = effectifsSheet.createRow(rowLine++);
		cell = outputRow.createCell(0);
		cell.setCellValue("Nom");
		cell = outputRow.createCell(1);
		cell.setCellValue("eMail");
		cell = outputRow.createCell(2);
		cell.setCellValue("Equipe");
		cell = outputRow.createCell(3);
		cell.setCellValue("Arbitrage Phase 1");
		cell = outputRow.createCell(4);
		cell.setCellValue("Table Phase 1");
		cell = outputRow.createCell(5);
		cell.setCellValue("Arbitrage Phase 2");
		cell = outputRow.createCell(6);
		cell.setCellValue("Table Phase 2");
		cell = outputRow.createCell(7);
		cell.setCellValue("Total");

		processRows(seniors, croissant, seniorsRecap);
		processRows(jeunes, decroissant, jeunesRecap);
		processRows(autres, croissant, autresRecap);

		Name namedRange = outputWb.createName();
		namedRange.setNameName("effectifs");
		namedRange.setRefersToFormula("'Effectifs'!$A$8:$A$" + rowLine);
		namedRange.setSheetIndex(outputWb.getSheetIndex(licenciesSheet));

		for (int i = 0; i < outputWb.getNumberOfSheets(); i++) {
			System.out.println(outputWb.getSheetName(i));
		}

		effectifsSheet.setColumnWidth(0, 25 * 256);
		for (int i = 1; i < 7; i++) {
			effectifsSheet.setColumnWidth(i, 20 * 256);
		}

	}

	private void processRows(HashMap<String, List<Licencie>> maps,
			Comparator<String> comparator, Row recapRow) {

		List<String> equipes = new ArrayList<>(maps.keySet());
		Collections.sort(equipes, comparator);

		StringBuffer phase1RecapFormula = new StringBuffer();
		StringBuffer phase2RecapFormula = new StringBuffer();

		for (String equipe : equipes) {
			Row equipeRow = effectifsSheet.createRow(rowLine++);
			Cell cell = equipeRow.createCell(0);
			cell.setCellValue(equipe + "________");

			List<Licencie> membres = maps.get(equipe);
			Collections.sort(membres, new ComparatorLicencie());
			for (Licencie membre : membres) {
				Row membreRow = effectifsSheet.createRow(rowLine++);
				cell = membreRow.createCell(0);
				cell.setCellValue(membre.nomPrenom);
				cell = membreRow.createCell(1);
				cell.setCellValue(membre.email);
				cell = membreRow.createCell(2);
				cell.setCellValue(equipe);
				String[] phases = new String[] { "Phase 1", "Phase 2" };
				int cellIdx = 3;
				for (String phase : phases) {
					membreRow.createCell(cellIdx).setCellFormula(
							"if(iserror(indirect(\"'" + phase
									+ "'!A1\")),0,countif(indirect(\"'" + phase
									+ "'!C:C\"),A"
									+ (membreRow.getRowNum() + 1)
									+ ") + countif(indirect(\"'" + phase
									+ "'!D:D\"),A"
									+ (membreRow.getRowNum() + 1) + "))");
					cellIdx++;
					membreRow.createCell(cellIdx).setCellFormula(
							"if(iserror(indirect(\"'" + phase
									+ "'!A1\")),0,countif(indirect(\"'" + phase
									+ "'!E:E\"),A"
									+ (membreRow.getRowNum() + 1)
									+ ") + countif(indirect(\"'" + phase
									+ "'!F:F\"),A"
									+ (membreRow.getRowNum() + 1) + "))");
					cellIdx++;
				}
				// Total
				membreRow.createCell(7).setCellFormula(
						"$D$" + (membreRow.getRowNum() + 1) + " + $E$"
								+ (membreRow.getRowNum() + 1) + " + $F$"
								+ (membreRow.getRowNum() + 1) + " + $G$"
								+ (membreRow.getRowNum() + 1));
			}

			// Arbitrage Phase 1
			equipeRow.createCell(3).setCellFormula(
					"sum($D$" + (equipeRow.getRowNum() + 2) + ":$D$" + rowLine
							+ ")");
			// Table Phase 1
			equipeRow.createCell(4).setCellFormula(
					"sum($E$" + (equipeRow.getRowNum() + 2) + ":$E$" + rowLine
							+ ")");
			// Arbitrage Phase 2
			equipeRow.createCell(5).setCellFormula(
					"sum($F$" + (equipeRow.getRowNum() + 2) + ":$F$" + rowLine
							+ ")");
			// Table Phase 2
			equipeRow.createCell(6).setCellFormula(
					"sum($G$" + (equipeRow.getRowNum() + 2) + ":$G$" + rowLine
							+ ")");
			// Total
			equipeRow.createCell(7).setCellFormula(
					"$D$" + (equipeRow.getRowNum() + 1) + " + $E$"
							+ (equipeRow.getRowNum() + 1) + " + $F$"
							+ (equipeRow.getRowNum() + 1) + " + $G$"
							+ (equipeRow.getRowNum() + 1));

			if (phase1RecapFormula.length() > 0) {
				phase1RecapFormula.append(" + ");
			}
			phase1RecapFormula.append("($D$" + (equipeRow.getRowNum() + 1)
					+ "+$E$" + (equipeRow.getRowNum() + 1) + ")");
			if (phase2RecapFormula.length() > 0) {
				phase2RecapFormula.append(" + ");
			}
			phase2RecapFormula.append("($F$" + (equipeRow.getRowNum() + 1)
					+ "+$G$" + (equipeRow.getRowNum() + 1) + ")");

		}

		recapRow.createCell(2).setCellFormula(phase1RecapFormula.toString());
		recapRow.createCell(3).setCellFormula(phase2RecapFormula.toString());
		recapRow.createCell(4).setCellFormula(
				"$C$" + (recapRow.getRowNum() + 1) + " + " + "$D$"
						+ (recapRow.getRowNum() + 1));

	}

}
