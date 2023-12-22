package com.excel.image.services.serviceImpl;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import com.excel.image.RequestBody.JsonToExcelConvertRequest;
import com.excel.image.services.ExcelService;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;

@Service
public class ExcelServiceImpl implements ExcelService {

	private ObjectMapper mapper = new ObjectMapper();

	@Override
	public String jsonToExcelConvertor(JsonToExcelConvertRequest jsonToExcelConvertRequest) {
		try {
			if (!jsonToExcelConvertRequest.getSourcePath().endsWith(".txt")) {
				throw new IllegalArgumentException("The source file should be .json file only");
			} else {
				Workbook workbook = null;
				File srcFile = new File(jsonToExcelConvertRequest.getSourcePath());
				// Creating workbook object based on target file format
				if (jsonToExcelConvertRequest.getDestination().endsWith(".xls")) {
					workbook = new HSSFWorkbook();
				} else if (jsonToExcelConvertRequest.getDestination().endsWith(".xlsx")) {
					workbook = new XSSFWorkbook();
				} else {
					throw new IllegalArgumentException("The target file extension should be .xls or .xlsx only");
				}

				// Reading the json file
				JsonNode jsonData = mapper.readTree(srcFile);

				// Iterating over the each sheets
				Iterator<String> sheetItr = jsonData.fieldNames();
//                System.out.println(sheetItr);
				int rowNum = 0;
				Sheet sheet = workbook.createSheet(jsonToExcelConvertRequest.getSheetName());
				while (sheetItr.hasNext()) {
					System.out.println(sheetItr);
					// create the workbook sheet
					String sheetName = sheetItr.next();

					ArrayNode sheetData = (ArrayNode) jsonData.get(sheetName);
					ArrayList<String> headers = createHeader(workbook, sheet, sheetData, rowNum);
					System.out.println(sheetData.size());
					System.out.println(rowNum);
					// Iterating over the each row data and writing into the sheet
					for (int i = 0; i < sheetData.size(); i++) {
						JsonNode rowData = sheetData.get(i);
						Row row = sheet.createRow(rowNum + 1);
						for (int j = 0; j < headers.size(); j++) {
							String value = rowData.get(headers.get(j)).asText();
							if (!value.isEmpty()) {
								if (value.charAt(0) == '{') {
									continue;
								}
							}
							row.createCell(j).setCellValue(value);
						}
						rowNum++;
					}
					rowNum++;
					/*
					 * automatic adjust data in column using autoSizeColumn, autoSizeColumn should
					 * be made after populating the data into the excel. Calling before populating
					 * data will not have any effect.
					 */
					for (int i = 0; i < headers.size(); i++) {
						sheet.autoSizeColumn(i);
					}

				}

				// creating a target file
//				String filename = srcFile.getName();
//				filename = filename.substring(0, filename.lastIndexOf(".json")) + targetFileExtension;
//				File targetFile = new File(srcFile.getParent(), filename);

				File targetFile = new File(jsonToExcelConvertRequest.getDestination());
				// write the workbook into target file
				FileOutputStream fos = new FileOutputStream(targetFile);
				workbook.write(fos);

				// close the workbook and fos
				workbook.close();
				fos.close();
				return "Successfully created Excel file at :" + jsonToExcelConvertRequest.getDestination();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return "Failed To Create Excel";
	}

	public ArrayList<String> createHeader(Workbook workbook, Sheet sheet, ArrayNode sheetData, int index) {

		ArrayList<String> headers = new ArrayList<>();
		// Creating cell style for header to make it bold
		CellStyle headerStyle = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setBold(true);
		headerStyle.setFont(font);

		// creating the header into the sheet
		Row header = sheet.createRow(index);
		Iterator<String> it = sheetData.get(0).fieldNames();
		int headerIdx = 0;
		while (it.hasNext()) {
			String headerName = it.next();
			headers.add(headerName);
			Cell cell = header.createCell(headerIdx++);
			cell.setCellValue(headerName);
			// apply the bold style to headers
			cell.setCellStyle(headerStyle);
		}
		return headers;
	}

	@Override
	public String excelToImageConvertor(JsonToExcelConvertRequest jsonToExcelConvertRequest) {
		// TODO Auto-generated method stub
		return null;
	}

}
