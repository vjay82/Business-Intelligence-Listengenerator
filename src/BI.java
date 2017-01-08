import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Random;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;

public class BI {
	protected static final LocalDate minDay = LocalDate.of(2014, 01, 01);
	protected static final LocalDate maxDay = LocalDate.of(2015, 12, 31);
	protected static final int ZIEL_LAGERBESTAND = 140;
	protected static final int MIN_LAGERBESTAND = 70;

	protected XSSFWorkbook wb;
	protected List<Integer> artikel = Lists.newArrayList();
	protected List<Integer> vertriebsregionen = Lists.newArrayList();
	protected List<Lager> lagerListe = Lists.newArrayList();
	protected Map<Integer, Integer> haendlerIdZuRegionId = Maps.newHashMap();
	protected Random random = new Random();

	protected class Lager {
		int vertriebsregionId;
		List<Integer> haendlerIds = Lists.newArrayList();
		Map<Integer, Integer> artikelIdAnzahl = Maps.newHashMap();
	}

	public BI() throws Exception {
		super();

		try (InputStream is = Files.newInputStream(Paths.get("BI_Listen.xlsx"))) {
			wb = new XSSFWorkbook(is);
		}

		// IDs einlesen
		readIdsToList("Artikel", artikel);
		readIdsToList("Vertriebsregion", vertriebsregionen);

		erzeugeHaendlerUndFachgeschaefte();

		// Lager erstellen
		erstelleLager();

		// Füllen der Anfangsbestände
		erstelleAnfangsbestaende();

		passTheYear();

		//		erstelleLagerEingangAusgangTabelle(); wird durch Importscript in Qlik Sense übernommen

		try (OutputStream os = Files.newOutputStream(Paths.get("BI_gefuellt.xlsx"))) {
			wb.write(os);
		}

		System.out.println("Done");

	}

	//	private void erstelleLagerEingangAusgangTabelle() {
	//		CellStyle cellStyle = wb.createCellStyle();
	//		CreationHelper createHelper = wb.getCreationHelper();
	//		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd.mm.yyyy"));
	//
	//		int index = 0;
	//		Date day1 = Date.from(LocalDate.of(2014, 01, 01).atStartOfDay(ZoneId.systemDefault()).toInstant());
	//		XSSFSheet sheetLager = wb.getSheet("Transaktionen");
	//
	//		// Anfangsbestände
	//		for (Lager lager : lagerListe) {
	//			for (int artikelId : artikel) {
	//				int anzahl = lager.artikelIdAnzahl.get(artikelId);
	//				createGetCell(sheetLager, ++index, 0).setCellValue(index);
	//				createGetCell(sheetLager, index, 1).setCellValue(day1);
	//				createGetCell(sheetLager, index, 1).setCellStyle(cellStyle);
	//				createGetCell(sheetLager, index, 2).setCellValue(artikelId);
	//
	//				createGetCell(sheetLager, index, 4).setCellValue(lager.vertriebsregionId);
	//				createGetCell(sheetLager, index, 6).setCellValue(anzahl);
	//			}
	//		}
	//
	//		// EK
	//		XSSFSheet ekSheet = wb.getSheet("EK");
	//		for (int ekIndex = 1; ekIndex <= ekSheet.getLastRowNum(); ekIndex++) {
	//			createGetCell(sheetLager, ++index, 0).setCellValue(index);
	//			createGetCell(sheetLager, index, 1).setCellValue(createGetCell(ekSheet, ekIndex, 1).getDateCellValue());
	//			createGetCell(sheetLager, index, 1).setCellStyle(cellStyle);
	//			createGetCell(sheetLager, index, 2).setCellValue(createGetCell(ekSheet, ekIndex, 3).getNumericCellValue());
	//
	//			createGetCell(sheetLager, index, 4).setCellValue(createGetCell(ekSheet, ekIndex, 2).getNumericCellValue());
	//			createGetCell(sheetLager, index, 6).setCellValue(createGetCell(ekSheet, ekIndex, 4).getNumericCellValue());
	//		}
	//
	//		// VK
	//		XSSFSheet vkSheet = wb.getSheet("VK");
	//		for (int vkIndex = 1; vkIndex <= vkSheet.getLastRowNum(); vkIndex++) {
	//			createGetCell(sheetLager, ++index, 0).setCellValue(index);
	//			createGetCell(sheetLager, index, 1).setCellValue(createGetCell(vkSheet, vkIndex, 1).getDateCellValue());
	//			createGetCell(sheetLager, index, 1).setCellStyle(cellStyle);
	//			createGetCell(sheetLager, index, 2).setCellValue(createGetCell(vkSheet, vkIndex, 2).getNumericCellValue());
	//			createGetCell(sheetLager, index, 3).setCellValue(createGetCell(vkSheet, vkIndex, 3).getNumericCellValue());
	//
	//			int haendlerId = (int) createGetCell(vkSheet, vkIndex, 3).getNumericCellValue();
	//			createGetCell(sheetLager, index, 4).setCellValue(haendlerIdZuRegionId.get(haendlerId));
	//
	//			createGetCell(sheetLager, index, 5).setCellValue(createGetCell(vkSheet, vkIndex, 4).getNumericCellValue());
	//		}
	//
	//		//		CellStyle cellStyle = wb.createCellStyle();
	//		//		CreationHelper createHelper = wb.getCreationHelper();
	//		//		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd.mm.yyyy"));
	//		//
	//		//		int index = 0;
	//		//		Date day1 = Date.from(LocalDate.of(2014, 01, 01).atStartOfDay(ZoneId.systemDefault()).toInstant());
	//		//		XSSFSheet sheetLager = wb.getSheet("LagerEingangAusgang");
	//		//
	//		//		// Anfangsbestände
	//		//		for (Lager lager : lagerListe) {
	//		//			for (int artikelId : artikel) {
	//		//				int anzahl = lager.artikelIdAnzahl.get(artikelId);
	//		//				createGetCell(sheetLager, ++index, 0).setCellValue(index);
	//		//				createGetCell(sheetLager, index, 1).setCellValue(day1);
	//		//				createGetCell(sheetLager, index, 1).setCellStyle(cellStyle);
	//		//				createGetCell(sheetLager, index, 2).setCellValue(lager.vertriebsregionId);
	//		//				createGetCell(sheetLager, index, 3).setCellValue(artikelId);
	//		//				createGetCell(sheetLager, index, 4).setCellValue(anzahl);
	//		//			}
	//		//		}
	//		//
	//		//		// EK
	//		//		XSSFSheet ekSheet = wb.getSheet("EK");
	//		//		for (int ekIndex = 1; ekIndex <= ekSheet.getLastRowNum(); ekIndex++) {
	//		//			createGetCell(sheetLager, ++index, 0).setCellValue(createGetCell(ekSheet, ekIndex, 0).getNumericCellValue());
	//		//			createGetCell(sheetLager, index, 1).setCellValue(createGetCell(ekSheet, ekIndex, 1).getDateCellValue());
	//		//			createGetCell(sheetLager, index, 1).setCellStyle(cellStyle);
	//		//			createGetCell(sheetLager, index, 2).setCellValue(createGetCell(ekSheet, ekIndex, 2).getNumericCellValue());
	//		//			createGetCell(sheetLager, index, 3).setCellValue(createGetCell(ekSheet, ekIndex, 3).getNumericCellValue());
	//		//			createGetCell(sheetLager, index, 4).setCellValue(createGetCell(ekSheet, ekIndex, 4).getNumericCellValue());
	//		//		}
	//		//
	//		//		// VK
	//		//		XSSFSheet vkSheet = wb.getSheet("VK");
	//		//		for (int vkIndex = 1; vkIndex <= vkSheet.getLastRowNum(); vkIndex++) {
	//		//			createGetCell(sheetLager, ++index, 0).setCellValue(createGetCell(vkSheet, vkIndex, 0).getNumericCellValue());
	//		//			createGetCell(sheetLager, index, 1).setCellValue(createGetCell(vkSheet, vkIndex, 1).getDateCellValue());
	//		//			createGetCell(sheetLager, index, 1).setCellStyle(cellStyle);
	//		//
	//		//			int haendlerId = (int) createGetCell(vkSheet, vkIndex, 3).getNumericCellValue();
	//		//			createGetCell(sheetLager, index, 2).setCellValue(haendlerIdZuRegionId.get(haendlerId));
	//		//
	//		//			createGetCell(sheetLager, index, 3).setCellValue(createGetCell(vkSheet, vkIndex, 2).getNumericCellValue());
	//		//			createGetCell(sheetLager, index, 4).setCellValue(createGetCell(vkSheet, vkIndex, 4).getNumericCellValue() * -1);
	//		//		}
	//	}

	private void erzeugeHaendlerUndFachgeschaefte() {
		XSSFSheet sheet = wb.getSheet("Haendler");

		int rowIndex = 0;
		int regionId = 1;
		for (int index = 0; index < 50 + 300; index++) {

			String type = index < 50 ? "Filiale" : "Fachgeschaeft";
			String name = type + " " + (index < 50 ? String.valueOf(index + 1) : String.valueOf(index - 50 + 1));

			if (index > 0 && index % 4 == 0) {
				regionId++;
				if (regionId > 4) {
					regionId = 1;
				}
			}

			createGetCell(sheet, ++rowIndex, 0).setCellValue(rowIndex);
			createGetCell(sheet, rowIndex, 1).setCellValue(name);
			createGetCell(sheet, rowIndex, 2).setCellValue(type);
			createGetCell(sheet, rowIndex, 3).setCellValue(index % 4 + 1);
			createGetCell(sheet, rowIndex, 4).setCellValue(regionId);

			haendlerIdZuRegionId.put(rowIndex, regionId);
		}

	}

	private void passTheYear() {
		int ekId = 1;
		int ekIndex = 0;
		int vkIndex = 0;
		int vkId = 1;

		XSSFSheet ekSheet = wb.getSheet("EK");
		XSSFSheet vkSheet = wb.getSheet("VK");

		LocalDate day = minDay;

		DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("dd.MM.yyyy");

		CellStyle cellStyle = wb.createCellStyle();
		CreationHelper createHelper = wb.getCreationHelper();
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd.mm.yyyy"));

		while (!day.isAfter(maxDay)) {
			if (!DayOfWeek.SUNDAY.equals(day.getDayOfWeek())) {

				String datumText = day.format(dateTimeFormatter);
				Date datumDate = Date.from(day.atStartOfDay(ZoneId.systemDefault()).toInstant());
				System.out.println("Berechne Tag: " + datumText);

				for (Lager lager : lagerListe) {

					// Verkauf für den Tag für das Lager
					for (int haendlerId : lager.haendlerIds) {
						for (int artikelId : artikel) {
							boolean wirdBestellt = random.nextInt(20) == 0; // 1/20 chance

							if (wirdBestellt) {

								// Es werden 1-10 Flaschen bestellt, es können maximal so viele geliefert werden wie noch vorrätig sind
								int bestand = lager.artikelIdAnzahl.get(artikelId);
								int anzahl = Math.min(random.nextInt(10 - 1) + 1, bestand);

								if (anzahl > 0) {
									lager.artikelIdAnzahl.put(artikelId, bestand - anzahl);
									createGetCell(vkSheet, ++vkIndex, 0).setCellValue(vkId++);
									createGetCell(vkSheet, vkIndex, 1).setCellValue(datumDate);
									createGetCell(vkSheet, vkIndex, 1).setCellStyle(cellStyle);
									createGetCell(vkSheet, vkIndex, 2).setCellValue(artikelId);
									createGetCell(vkSheet, vkIndex, 3).setCellValue(haendlerId);
									createGetCell(vkSheet, vkIndex, 4).setCellValue(anzahl);
								} else {
									System.out.println("Lager leer, artikelId: " + artikelId);
								}

							}

						}

					}

					// Am Freitag wird neue Ware geordert um das Lager aufzufüllen.
					// Diese trifft dann immer Montags ein.
					if (!DayOfWeek.MONDAY.equals(day.getDayOfWeek())) {
						for (int artikelId : artikel) {
							int bestand = lager.artikelIdAnzahl.get(artikelId);
							if (bestand < MIN_LAGERBESTAND) {
								int anzahl = ZIEL_LAGERBESTAND - bestand;
								lager.artikelIdAnzahl.put(artikelId, ZIEL_LAGERBESTAND);
								createGetCell(ekSheet, ++ekIndex, 0).setCellValue(ekId++);
								createGetCell(ekSheet, ekIndex, 1).setCellValue(datumDate);
								createGetCell(ekSheet, ekIndex, 1).setCellStyle(cellStyle);
								createGetCell(ekSheet, ekIndex, 2).setCellValue(lager.vertriebsregionId);
								createGetCell(ekSheet, ekIndex, 3).setCellValue(artikelId);
								createGetCell(ekSheet, ekIndex, 4).setCellValue(anzahl);
							}
						}
					}

				}
			}

			day = day.plusDays(1);
		}

	}

	private Cell createGetCell(XSSFSheet sheet, int rowIndex, int columnIndex) {
		Row row = sheet.getRow(rowIndex);
		if (row == null) {
			row = sheet.createRow(rowIndex);
		}

		Cell cell = row.getCell(columnIndex);
		if (cell == null) {
			cell = row.createCell(columnIndex);
		}

		return cell;
	}

	private void erstelleAnfangsbestaende() {
		CellStyle cellStyle = wb.createCellStyle();
		CreationHelper createHelper = wb.getCreationHelper();
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd.mm.yyyy"));
		Date ersterErster = Date.from(LocalDate.of(2014, 01, 01).atStartOfDay(ZoneId.systemDefault()).toInstant());
		XSSFSheet sheet = wb.getSheet("Lager Anfangsbestaende");
		int index = 0;
		for (Lager lager : lagerListe) {
			for (int artikelId : artikel) {
				int anzahl = Math.max(MIN_LAGERBESTAND, Math.min(ZIEL_LAGERBESTAND, random.nextInt(ZIEL_LAGERBESTAND * 2))); // ich möchte dass bei gleichverteilung mehr lager gefüllt sind

				lager.artikelIdAnzahl.put(artikelId, anzahl);
				createGetCell(sheet, ++index, 0).setCellValue(ersterErster);
				createGetCell(sheet, index, 0).setCellStyle(cellStyle);
				createGetCell(sheet, index, 1).setCellValue(lager.vertriebsregionId);
				createGetCell(sheet, index, 2).setCellValue(artikelId);
				createGetCell(sheet, index, 3).setCellValue(anzahl);
			}

		}

		//		CellStyle cellStyle = wb.createCellStyle();
		//		CreationHelper createHelper = wb.getCreationHelper();
		//		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd.mm.yyyy"));
		//		LocalDate day = LocalDate.of(2014, 01, 01);
		//		Date datumDate = Date.from(day.atStartOfDay(ZoneId.systemDefault()).toInstant());
		//		XSSFSheet ekSheet = wb.getSheet("EK");
		//		int index = 0;
		//		for (Lager lager : lagerListe) {
		//			for (int artikelId : artikel) {
		//				int anzahl = Math.max(MIN_LAGERBESTAND, Math.min(ZIEL_LAGERBESTAND, random.nextInt(ZIEL_LAGERBESTAND * 2))); // ich möchte dass bei gleichverteilung mehr lager gefüllt sind
		//
		//				lager.artikelIdAnzahl.put(artikelId, anzahl);
		//				createGetCell(ekSheet, ++ekIndex, 0).setCellValue(ekId++);
		//				createGetCell(ekSheet, ekIndex, 1).setCellValue(datumDate);
		//				createGetCell(ekSheet, ekIndex, 1).setCellStyle(cellStyle);
		//				createGetCell(ekSheet, ekIndex, 2).setCellValue(lager.vertriebsregionId);
		//				createGetCell(ekSheet, ekIndex, 3).setCellValue(artikelId);
		//				createGetCell(ekSheet, ekIndex, 4).setCellValue(anzahl);
		//			}
		//
		//		}
	}

	private void erstelleLager() {
		for (int vertriebsregionId : vertriebsregionen) {
			Lager lager = new Lager();
			lagerListe.add(lager);
			lager.vertriebsregionId = vertriebsregionId;
			System.out.println("Erstelle lager für region " + vertriebsregionId);

			// Haendler einfügen
			XSSFSheet sheet = wb.getSheet("Haendler");
			Iterator<Row> rowIterator = sheet.rowIterator();
			Row headerRow = rowIterator.next();
			int vertriebsRegionColumnIndex = -1;
			Iterator<Cell> cellIterator = headerRow.cellIterator();
			while (cellIterator.hasNext()) {
				vertriebsRegionColumnIndex++;
				if ("Vertriebsregion_ID".equals(cellIterator.next().getStringCellValue())) {
					break;
				}
			}

			int finalVertriebsRegionColumnIndex = vertriebsRegionColumnIndex;
			rowIterator.forEachRemaining(row -> {
				int haendlerVertriebsregionId = (int) row.getCell(finalVertriebsRegionColumnIndex).getNumericCellValue();
				if (haendlerVertriebsregionId == vertriebsregionId) {
					int id = (int) row.getCell(0).getNumericCellValue();
					lager.haendlerIds.add(id);
					System.out.println("Lager: " + vertriebsregionId + " fuege haendler " + id + " hinzu.");
				}
			});
		}
	}

	private void readIdsToList(String sheetName, List<Integer> list) {
		Iterator<Row> rowIterator = wb.getSheet(sheetName).rowIterator();
		rowIterator.next();
		rowIterator.forEachRemaining(row -> {
			int id = (int) row.getCell(0).getNumericCellValue();
			list.add(id);
			System.out.println(sheetName + " id: " + id);
		});
	}

	public static void main(String[] args) throws Exception {
		new BI();
	}

}
