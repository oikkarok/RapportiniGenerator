package oikkarok.GenarateExcel.controller;

import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import oikkarok.GenarateExcel.service.GenarateExcelFromRapportiniService;

@RestController
public class GenarateExcelFromRapportiniController {

	@Autowired
	GenarateExcelFromRapportiniService service;

	@PostMapping("/generateExcel")
	public ResponseEntity<byte[]> generateExcelFromHtml(@RequestBody String htmlString) {

		try {
			List<List<String>> tableData = service.extractTableData(htmlString);
			Workbook workbook = service.createWorkbookWithData(tableData);
			byte[] excelBytes = service.createExcelBytes(workbook);
			service.writeWorkbookToFile(workbook);
			HttpHeaders headers = service.createHttpHeaders(excelBytes);

			return new ResponseEntity<>(excelBytes, headers, HttpStatus.OK);

		} catch (IOException e) {
			e.printStackTrace();
			return new ResponseEntity<>("Error generating Excel file".getBytes(), HttpStatus.INTERNAL_SERVER_ERROR);
		}
	}

}
