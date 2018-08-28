package bcbl.planning.arbitrage;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.List;
import java.util.Locale;
import java.util.Properties;

import org.apache.commons.lang3.text.WordUtils;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellUtil;

public class CreationOngletPhase {

	public final static String BCBL = "BASKET CLUB BASSE LOIRE";
	public final static String ENTENTE = "LOIRE ACHENEAU BASKET";
	public final static String EXEMPT = "Exempt";

	public final static int DIVISION_COLIDX = 0;
	public final static int EQUIPE1_COLIDX = 2;
	public final static int EQUIPE2_COLIDX = 3;
	public final static int DATE_RENCONTRE_COLIDX = 4;
	public final static int HEURE_RENCONTRE_COLIDX = 5;
	public final static int SALLE_COLIDX = 6;

	private HSSFSheet extractFbiSheet;
	private Workbook outputWb;
	private Sheet ongletSheet;
	private Properties mappings;

	private int ongletRowIndex = 0;
	private CellStyle headerCS;
	private CellStyle contentCS;
	private CellStyle tableRowCS;
	private CellStyle tableBottomCS;
	private CellStyle tableHeaderCS;
	private CellStyle col0CS;
	private Font femininFont;
	private Font masculinFont;
	private Font exterieurFont;

	private static String convertTeamToClub(String team) {
		// Pour ne pas avoir les indices d'équipe Comme par exemple AL Saint
		// Sebastien sur Loire - 2
		int lastTiret = team.lastIndexOf("-");
		if (lastTiret >= 0) {
			try {
				Integer.parseInt(team.substring(lastTiret + 1).trim());
				// dans ce cas, il y a bien un indice d'équipe, on le retire
				// pour n'avoir que le club
				return team.substring(0, lastTiret).trim();
			} catch (Exception e) {
			}
		}
		return team;
	}

	public CreationOngletPhase(HSSFSheet extractFbiSheet, Workbook outputWb, Sheet ongletSheet, Properties mappings) {
		this.extractFbiSheet = extractFbiSheet;
		this.outputWb = outputWb;
		this.ongletSheet = ongletSheet;
		this.mappings = mappings;
	}

	private CellStyle getHeaderCS() {
		if (headerCS == null) {
			headerCS = outputWb.createCellStyle();
			Font font = outputWb.createFont();
			font.setFontName("Calibri");
			font.setFontHeightInPoints((short) 13);
			headerCS.setFont(font);
		}
		return headerCS;
	}

	private CellStyle getContentCS() {
		if (contentCS == null) {
			contentCS = outputWb.createCellStyle();
			Font font = outputWb.createFont();
			font.setFontName("Calibri");
			font.setFontHeightInPoints((short) 11);
			contentCS.setFont(font);
		}
		return contentCS;
	}

	private CellStyle getCol0CS() {
		if (col0CS == null) {
			col0CS = outputWb.createCellStyle();
			col0CS.cloneStyleFrom(getContentCS());
			col0CS.setDataFormat(outputWb.createDataFormat().getFormat("ddd. hh:mm"));
		}
		return col0CS;
	}

	private CellStyle getTableRowCS() {
		if (tableRowCS == null) {
			tableRowCS = outputWb.createCellStyle();
			tableRowCS.cloneStyleFrom(getContentCS());
			tableRowCS.setAlignment(CellStyle.ALIGN_CENTER);
			tableRowCS.setBorderBottom(CellStyle.BORDER_THIN);
			tableRowCS.setBorderLeft(CellStyle.BORDER_MEDIUM);
			tableRowCS.setBorderRight(CellStyle.BORDER_MEDIUM);
		}
		return tableRowCS;
	}

	private CellStyle getTableBottomCS() {
		if (tableBottomCS == null) {
			tableBottomCS = outputWb.createCellStyle();
			tableBottomCS.cloneStyleFrom(getTableRowCS());
			tableBottomCS.setBorderBottom(CellStyle.BORDER_MEDIUM);
		}
		return tableBottomCS;
	}

	public void doIt() throws IOException {
		int rows = extractFbiSheet.getPhysicalNumberOfRows();

		System.out.println("physicalNumberOfRows = " + rows);
		Calendar currentWeek = null;
		List<HSSFRow> rowsParWeekend = new ArrayList<HSSFRow>();

		CellRangeAddressList addressList = new CellRangeAddressList();

		// La 1ere ligne contient les headers. Les lignes sont triés
		// chronologiquement ascendant
		for (int r = 1; r < rows; r++) {
			HSSFRow row = extractFbiSheet.getRow(r);
			if (row != null && (row.getCell(EQUIPE1_COLIDX) != null)) {

				String equipe1 = row.getCell(EQUIPE1_COLIDX).getStringCellValue();
				String equipe2 = row.getCell(EQUIPE2_COLIDX).getStringCellValue();
				if (equipe1.contains(BCBL) || equipe1.contains(ENTENTE) || equipe2.contains(BCBL)
						|| equipe2.contains(ENTENTE)) {

					HSSFCell cell = row.getCell(DATE_RENCONTRE_COLIDX);

					if (currentWeek == null) {
						currentWeek = GregorianCalendar.getInstance(Locale.FRANCE);
						currentWeek.setTime(cell.getDateCellValue());
						rowsParWeekend.add(row);
					} else {
						// Si méme semaine, on conserve dans le méme set
						Calendar calendar = GregorianCalendar.getInstance(Locale.FRANCE);
						calendar.setTime(cell.getDateCellValue());
						if (calendar.get(Calendar.WEEK_OF_YEAR) == currentWeek.get(Calendar.WEEK_OF_YEAR)) {
							rowsParWeekend.add(row);
						} else {

							processRows(currentWeek, rowsParWeekend, addressList);

							// Nouveau week-end
							currentWeek = calendar;
							rowsParWeekend.clear();
							rowsParWeekend.add(row);
						}
					}
				}
			}
		}

		if (!rowsParWeekend.isEmpty()) {
			processRows(currentWeek, rowsParWeekend, addressList);
			rowsParWeekend.clear();
		}

		DVConstraint dvConstraint = DVConstraint.createFormulaListConstraint("effectifs");
		DataValidation dataValidation = new HSSFDataValidation(addressList, dvConstraint);
		dataValidation.setSuppressDropDownArrow(false);
		dataValidation.setEmptyCellAllowed(true);

		ongletSheet.addValidationData(dataValidation);

		ongletSheet.setColumnWidth(0, 12 * 256);
		ongletSheet.setColumnWidth(1, 12 * 256);
		for (int i = 2; i < 7; i++) {
			ongletSheet.setColumnWidth(i, 20 * 256);
		}

	}

	private void processRows(Calendar weekend, List<HSSFRow> rows, CellRangeAddressList addressList) {
		Date samedi;
		Date dimanche;
		if (weekend.get(Calendar.DAY_OF_WEEK) == Calendar.SATURDAY) {
			samedi = weekend.getTime();
			weekend.add(Calendar.DAY_OF_YEAR, 1);
			dimanche = weekend.getTime();
		} else if (weekend.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY) {
			dimanche = weekend.getTime();
			weekend.add(Calendar.DAY_OF_YEAR, -1);
			samedi = weekend.getTime();
		} else {
			System.out.println("Probleme de dates");
			return;
		}

		DateFormat df1 = new SimpleDateFormat("dd MMMM", Locale.FRANCE);
		DateFormat df2 = new SimpleDateFormat("dd MMMM yyyy", Locale.FRANCE);

		Row outputRow = ongletSheet.createRow(ongletRowIndex++);
		createHeaderCell(outputRow, 0, df1.format(samedi) + " au " + df2.format(dimanche));

		SheetConditionalFormatting sheetCF = ongletSheet.getSheetConditionalFormatting();

		ongletRowIndex++;
		outputRow = ongletSheet.createRow(ongletRowIndex++);
		createTableHeaderCells(outputRow,
				new String[] { "Catégorie", "1er Arbitre", "2ème Arbitre", "Marqueur", "Chronométreur" });

		Collections.sort(rows, new Comparator<HSSFRow>() {
			public int compare(HSSFRow o1, HSSFRow o2) {
				String equipe1 = getNomEquipe(o1);
				String equipe2 = getNomEquipe(o2);

				if (equipe1.startsWith("U")) {
					if (equipe2.startsWith("U")) {
						return equipe1.compareTo(equipe2);
					} else {
						return -1;
					}
				} else if (equipe1.startsWith("R")) {
					if (equipe2.startsWith("U")) {
						return 1;
					} else if (equipe2.startsWith("D")) {
						return -1;
					}
				} else if (equipe1.startsWith("D")) {
					if (equipe2.startsWith("U")) {
						return 1;
					} else if (equipe2.startsWith("R")) {
						return 1;
					}
				}
				return equipe1.compareTo(equipe2);
			}
		});

		for (int i = 0; i < rows.size(); i++) {
			HSSFRow row = rows.get(i);
			outputRow = ongletSheet.createRow(ongletRowIndex++);

			//
			boolean exempt = row.getCell(EQUIPE1_COLIDX).getStringCellValue().contains(EXEMPT)
					|| row.getCell(EQUIPE2_COLIDX).getStringCellValue().contains(EXEMPT);

			String equipeRecevant = row.getCell(EQUIPE1_COLIDX).getStringCellValue();
			boolean domicile = equipeRecevant.contains(BCBL) || equipeRecevant.contains(ENTENTE);

			if (exempt) {
				outputRow.createCell(0).setCellValue("Exempt");
			} else {
				GregorianCalendar gcMatch = new GregorianCalendar();
				gcMatch.setTime(row.getCell(DATE_RENCONTRE_COLIDX).getDateCellValue());
				
				GregorianCalendar gcHeure = new GregorianCalendar(); 
				gcHeure.setTime(row.getCell(HEURE_RENCONTRE_COLIDX).getDateCellValue());
				
				gcMatch.set(Calendar.HOUR_OF_DAY, gcHeure.get(Calendar.HOUR_OF_DAY));
				gcMatch.set(Calendar.MINUTE, gcHeure.get(Calendar.MINUTE));
				gcMatch.set(Calendar.SECOND, 0);
				gcMatch.set(Calendar.MILLISECOND, 0);
								
				outputRow.createCell(0).setCellValue(gcMatch.getTime());				
			}
			outputRow.getCell(0).setCellStyle(getCol0CS());
			if (exempt || !domicile) {
				CellUtil.setFont(outputRow.getCell(0), outputWb, getExterieurFont());
			}
			
			// 
			if (!exempt && domicile) {
				String salle = row.getCell(SALLE_COLIDX).getStringCellValue();
				if (salle.toLowerCase().contains("rigaudeau")) {
					outputRow.createCell(1).setCellValue("Rigaudeau");					
				} else if (salle.toLowerCase().contains("flute")) {
					outputRow.createCell(1).setCellValue("Flûte");					
				} else if (salle.toLowerCase().contains("bellestre")) {
					outputRow.createCell(1).setCellValue("Bouaye");
				} else {
					outputRow.createCell(1).setCellValue("Souvré");
				}
			}

			//
			String nomEquipeBcbl = getNomEquipe(row);
			String[] values = new String[5];
			values[0] = nomEquipeBcbl;
			// Officiels
			boolean officiels = (nomEquipeBcbl.startsWith("R") || nomEquipeBcbl.contains("Région"))
					&& (!nomEquipeBcbl.contains("Réserve"));

			if (nomEquipeBcbl.equals("DF1")) {
				officiels = true;
			}

			if (domicile && officiels) {
				values[1] = values[2] = "OFFICIELS";
			}
			if (i < rows.size() - 1) {
				createTableRowCells(outputRow, values);
			} else {
				createTableBottomCells(outputRow, values);
			}

			// Feminin
			boolean feminin = nomEquipeBcbl.contains("F");
			if (feminin) {
				CellUtil.setFont(outputRow.getCell(2), outputWb, getFemininFont());
			} else {
				CellUtil.setFont(outputRow.getCell(2), outputWb, getMasculinFont());
			}

			if (exempt || !domicile) {
				for (int j = 1; j < values.length; j++) {
					CellUtil.setCellStyleProperty(outputRow.getCell(j + 2), outputWb, CellUtil.FILL_PATTERN,
							CellStyle.THIN_FORWARD_DIAG);
				}
			}

			//
			// Pour chacune des cellules non vides, on rajoute une
			// validation
			if (!exempt) {
				if (domicile) {
					outputRow.createCell(7).setCellValue("Domicile");
					outputRow.getCell(7).setCellStyle(getContentCS());
					outputRow.createCell(8).setCellValue(WordUtils
							.capitalizeFully(convertTeamToClub(row.getCell(EQUIPE2_COLIDX).getStringCellValue())));
					outputRow.getCell(8).setCellStyle(getContentCS());

					CellRangeAddress cra;
					int rowIndex = ongletRowIndex - 1;
					if (domicile && officiels) {
						cra = new CellRangeAddress(rowIndex, rowIndex, 4, 5);
					} else {
						cra = new CellRangeAddress(rowIndex, rowIndex, 2, 5);
					}
					addressList.addCellRangeAddress(cra);

					// Conditional Formatting
					if (false) {
						String formula = "AND(ISERROR(FIND(\"sam\",A" + ongletRowIndex + ",1)),ISERROR(FIND(\"dim\",A"
								+ ongletRowIndex + ",1)))";
						ConditionalFormattingRule cfr = sheetCF.createConditionalFormattingRule(formula);
						PatternFormatting pf = cfr.createPatternFormatting();
						pf.setFillPattern(CellStyle.THIN_FORWARD_DIAG);

						sheetCF.addConditionalFormatting(new CellRangeAddress[] { cra }, cfr);
					}
				} else {
					// club recevant
					outputRow.createCell(7).setCellValue("Exterieur");
					CellUtil.setFont(outputRow.getCell(7), outputWb, getExterieurFont());

					outputRow.createCell(8).setCellValue(WordUtils.capitalizeFully(convertTeamToClub(equipeRecevant)));
					CellUtil.setFont(outputRow.getCell(8), outputWb, getExterieurFont());
				}
			}
		}

		outputRow = ongletSheet.createRow(ongletRowIndex++);

	}

	private Font getExterieurFont() {
		if (exterieurFont == null) {
			exterieurFont = outputWb.createFont();
			exterieurFont.setFontName("Calibri");
			exterieurFont.setFontHeightInPoints((short) 11);
			exterieurFont.setColor(HSSFColor.GREY_25_PERCENT.index);
		}
		return exterieurFont;
	}

	private Font getFemininFont() {
		if (femininFont == null) {
			femininFont = outputWb.createFont();
			femininFont.setFontName("Calibri");
			femininFont.setFontHeightInPoints((short) 11);
			femininFont.setColor(HSSFColor.PINK.index);
		}
		return femininFont;
	}

	private Font getMasculinFont() {
		if (masculinFont == null) {
			masculinFont = outputWb.createFont();
			masculinFont.setFontName("Calibri");
			masculinFont.setFontHeightInPoints((short) 11);
			masculinFont.setColor(HSSFColor.BLUE.index);
		}
		return masculinFont;
	}

	private void createTableHeaderCells(Row outputRow, String[] headers) {
		createTableRowCells(outputRow, headers, getTableHeaderCS());
	}

	private CellStyle getTableHeaderCS() {
		if (tableHeaderCS == null) {
			tableHeaderCS = outputWb.createCellStyle();
			tableHeaderCS.cloneStyleFrom(getTableRowCS());
			tableHeaderCS.setBorderTop(CellStyle.BORDER_MEDIUM);
			tableHeaderCS.setBorderBottom(CellStyle.BORDER_MEDIUM);
		}
		return tableHeaderCS;
	}

	private void createTableBottomCells(Row outputRow, String[] values) {
		createTableRowCells(outputRow, values, getTableBottomCS());
	}

	private void createTableRowCells(Row outputRow, String[] values) {
		createTableRowCells(outputRow, values, getTableRowCS());
	}

	private void createTableRowCells(Row outputRow, String[] values, CellStyle cs) {
		for (int i = 0; i < values.length; i++) {
			Cell cell = outputRow.createCell(i + 2);
			cell.setCellValue(values[i]);
			cell.setCellStyle(cs);
		}
	}

	private void createHeaderCell(Row outputRow, int i, String value) {
		Cell header = outputRow.createCell(i);
		header.setCellValue(value);
		header.setCellStyle(getHeaderCS());
	}

	private String getNomEquipe(HSSFRow row) {
		String division = row.getCell(DIVISION_COLIDX).getStringCellValue();
		String equipe = row.getCell(EQUIPE1_COLIDX).getStringCellValue();
		if (!equipe.contains(BCBL) && !equipe.contains(ENTENTE)) {
			equipe = row.getCell(EQUIPE2_COLIDX).getStringCellValue();
		}
		// Regarder dans la table de correspondance avec division et nom
		// d'équipe
		String result = mappings.getProperty(division + "." + equipe);
		if (result == null) {

			if (division.startsWith("U")) {
				// Si commence par U, c'est une equipe jeune département
				int tiret = division.indexOf("-");
				if (tiret > 0) {
					result = division.substring(0, tiret);
				} else if (division.toUpperCase().endsWith("1PH") || division.toUpperCase().endsWith("1PH")) {
					result = division.substring(0, division.length() - 3);
				}
				String suffixe = "";
				int index = equipe.indexOf(ENTENTE);
				if (index >= 0) {
					suffixe = equipe.substring(index + ENTENTE.length()) + " (CTC)";
				} else {
					index = equipe.indexOf(BCBL);
					if (index >= 0) {
						suffixe = equipe.substring(index + BCBL.length());
					}
				}
				result = result + suffixe;
			} else if (division.startsWith("R") && division.charAt(2) == 'U') {
				// Si commence par R et U en 2eme caractére, c'est une équipe
				// jeune région
				result = division.substring(2) + division.charAt(1) + " - 1 Région";
			} else {
				// Sinon, équipe logique
				result = division;
			}
		}
		return result;
	}

}
