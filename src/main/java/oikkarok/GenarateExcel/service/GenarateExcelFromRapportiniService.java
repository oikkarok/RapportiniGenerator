package oikkarok.GenarateExcel.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Service;

import oikkarok.GenarateExcel.entities.Progetto;
import oikkarok.GenarateExcel.entities.Rapportino;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service
public class GenarateExcelFromRapportiniService {

	public static final String filePath = "C:\\Users\\s.coraccio\\Projects\\GenarateExcel\\download\\generated_excel.xlsx";

	private static int rowIndex;

	/**
	 * 
	 * @param workbook
	 * @return
	 * @throws IOException
	 */
	public byte[] createExcelBytes(Workbook workbook) throws IOException {

		byte[] excelBytes;
		try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
			workbook.write(outputStream);
			excelBytes = outputStream.toByteArray();
		}

		return excelBytes;
	}

	/**
	 * 
	 * @param excelBytes
	 * @return
	 */
	public HttpHeaders createHttpHeaders(byte[] excelBytes) {

		HttpHeaders headers = new HttpHeaders();
		headers.setContentType(
				MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
		headers.setContentDispositionFormData("attachment", "generated_excel.xlsx");
		headers.setContentLength(excelBytes.length);

		return headers;
	}

	/**
	 * 
	 * @param workbook
	 * @throws IOException
	 */
	public void writeWorkbookToFile(Workbook workbook) throws IOException {
		FileOutputStream fileOutputStream = new FileOutputStream(new File(filePath));
		workbook.write(fileOutputStream);
		fileOutputStream.close();
		workbook.close();
	}

	/**
	 * 
	 * @param htmlString
	 * @return
	 */
	public List<List<String>> extractTableData(String htmlString) {

		List<List<String>> tableData = new ArrayList<>();
		Document doc = Jsoup.parse(htmlString);
		Element table = doc.select("table").first();

		if (table != null) {
			Elements rows = table.select("tr");
			for (Element row : rows) {
				Elements cells = row.select("td");
				List<String> rowData = new ArrayList<>();
				for (Element cell : cells) {
					rowData.add(cell.text());
				}
				if (!rowData.isEmpty()) {
					tableData.add(rowData);
				}
			}
		}

		return tableData;
	}

	/**
	 * 
	 * @param data
	 * @return
	 * @throws IOException
	 */
	public Workbook createWorkbookWithData(List<List<String>> data) throws IOException {

		Workbook workbook;
		Sheet sheet;
		File excel = new File(filePath);

		if (excel.exists() && excel.length() > 0) {
			FileInputStream fileInputStream = new FileInputStream(excel);
			workbook = new XSSFWorkbook(fileInputStream);
			sheet = workbook.getSheetAt(0); // Assume che il foglio di lavoro sia il primo (indice 0)
			rowIndex = sheet.getLastRowNum();
		} else {
			workbook = new XSSFWorkbook();
			sheet = workbook.createSheet();
		}

		buildRapportini(data, sheet);

		return workbook;
	}

	/**
	 * 
	 * @param data
	 * @param sheet
	 */
	private void buildRapportini(List<List<String>> data, Sheet sheet) {

		for (List<String> rowData : data) {
			List<String> firstRowData = rowData.subList(0, 5);
			List<String> secondRowData = rowData.subList(5, rowData.size());

			String[] splitFirstRowData = rowData.get(4).split("-");
			String[] splitSecondRowData = rowData.get(9).split("-");

			if (!splitFirstRowData[0].equals(rowData.get(4)) && splitFirstRowData.length > 0
					&& splitFirstRowData[0] != "") {

				for (String splitData : splitFirstRowData) {
					List<String> newRowData = new ArrayList<>(firstRowData);
					newRowData.set(3, splitData.trim().substring(splitData.trim().length() - 2));
					newRowData.set(4, splitData.trim().substring(0, splitData.length() - 3));
					createRapportinoRow(sheet, newRowData);
				}
			} else {
				createRapportinoRow(sheet, firstRowData);
			}

			if (!splitSecondRowData[0].equals(rowData.get(9)) && splitSecondRowData.length > 0
					&& splitSecondRowData[0] != "") {

				for (String splitData : splitSecondRowData) {
					List<String> newRowData = new ArrayList<>(secondRowData);
					newRowData.set(4, splitData.trim().substring(splitData.trim().length() - 2));
					newRowData.set(3, splitData.trim().substring(0, splitData.length() - 3));
					createRapportinoRow(sheet, newRowData);
				}
			} else {
				createRapportinoRow(sheet, secondRowData);
			}
		}
	}

	/**
	 * 
	 * @param sheet
	 * @param rowData
	 */
	private void createRapportinoRow(Sheet sheet, List<String> rowData) {

		Rapportino rapportino = new Rapportino();
		rapportino.setData(rowData.get(0) != "" ? rowData.get(0).substring(0, 5) : null);
		rapportino.setDurataEffettiva(rowData.get(3) != "" ? Integer.parseInt(rowData.get(3).trim()) : 0);
		rapportino.setDurataStimata(rowData.get(3) != "" ? Integer.parseInt(rowData.get(3).trim()) : 0);
		rapportino.setDescrizione(rowData.size() > 4 ? rowData.get(4) : null);

		rapportino.setNomeProgetto(rowData.size() > 4 && rapportino.getDescrizione() != ""
				? Progetto.valueOf(rowData.get(4).substring(0, 3))
				: null);

		rapportino.setResponsabile(rapportino.getNomeProgetto() != null
				? Rapportino.getMappaResponsabili().get(rapportino.getNomeProgetto())
				: null); // Ottieni il responsabile associato al progetto

		if (isComplete(rapportino)) {
			rowIndex++;
			Row row = sheet.createRow(rowIndex);
			row.createCell(0).setCellValue(Rapportino.proprietario);
			row.createCell(1).setCellValue(rapportino.getData().toString());
			row.createCell(2).setCellValue(rapportino.getDurataStimata());
			row.createCell(3).setCellValue(rapportino.getDurataEffettiva());
			row.createCell(4).setCellValue(rapportino.getNomeProgetto().toString());
			row.createCell(5).setCellValue(rapportino.getDescrizione());
			row.createCell(6).setCellValue(rapportino.getResponsabile().toString());
		}
	}

	/**
	 * 
	 * @param rapportino
	 * @return
	 */
	private boolean isComplete(Rapportino rapportino) {

		if (rapportino.getData() != null && rapportino.getDurataStimata() != 0 && rapportino.getDurataEffettiva() != 0
				&& rapportino.getNomeProgetto() != null && rapportino.getDescrizione() != null
				&& rapportino.getResponsabile() != null) {
			return true;
		}

		return false;
	}

}
