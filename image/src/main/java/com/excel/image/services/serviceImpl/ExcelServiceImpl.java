package com.excel.image.services.serviceImpl;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.imageio.ImageIO;

import org.apache.batik.transcoder.Transcoder;
import org.apache.batik.transcoder.TranscoderException;
import org.apache.batik.transcoder.TranscoderInput;
import org.apache.batik.transcoder.TranscoderOutput;
import org.apache.batik.transcoder.image.PNGTranscoder;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
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
			return "Failed To Create Excel. Cause : "+e.getLocalizedMessage();
		}
		
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
		/*
		 * try { System.out.println(jsonToExcelConvertRequest.getSourcePath());
		 * FileInputStream fis = new FileInputStream(new
		 * File(jsonToExcelConvertRequest.getSourcePath())); Workbook workbook = new
		 * XSSFWorkbook(fis); Sheet sheet = workbook.getSheetAt(0);
		 * 
		 * BufferedImage image = new BufferedImage(sheet.getColumnWidth(0),
		 * sheet.getLastRowNum() * sheet.getDefaultRowHeight(),
		 * BufferedImage.TYPE_INT_RGB);
		 * 
		 * System.out.println(image); Graphics2D graphics = image.createGraphics();
		 * graphics.setBackground(java.awt.Color.black); graphics.drawRect(0, 0,
		 * image.getWidth(), image.getHeight());
		 * 
		 * for (Row row : sheet) { for (Cell cell : row) {
		 * graphics.drawString(cell.toString(), (int) cell.getColumnIndex() * 100, (int)
		 * row.getRowNum() * 20); } }
		 * 
		 * String imageFilePath = "D:\\sts workspace\\excelToImage.png";
		 * ImageIO.write(image, "png", new File(imageFilePath));
		 * 
		 * graphics.dispose(); workbook.close(); fis.close();
		 * 
		 * return "Excel converted to image successfully!"; } catch (IOException e) {
		 * e.printStackTrace(); return "Error converting Excel to image."; } }
		 * 
		 * @PostMapping("/excel-to-image") public boolean
		 * convertExcelToImage1(@RequestBody JsonToExcelConvertRequest
		 * jsonToExcelConvertRequest) throws IOException, TranscoderException {
		 */
		try {
			FileInputStream fis = new FileInputStream(new File(jsonToExcelConvertRequest.getSourcePath()));
			Workbook workbook = new XSSFWorkbook(fis);
			Sheet sheet = workbook.getSheetAt(0);

			// Create SVG content from Excel data
			String svgContent = generateSvgFromExcel(workbook, sheet);

			// Convert SVG content to an image using Apache Batik
			byte[] imageBytes = convertSvgToImage(svgContent);
			System.out.println(imageBytes);
			BufferedImage bufferedImage = bytesToBufferedImage(imageBytes);
			if (ImageIO.write(bufferedImage, "png", new File(jsonToExcelConvertRequest.getDestination() + ".png"))) {
				return "Image Created Successfully at : " + jsonToExcelConvertRequest.getDestination() + ".png";
			} else {
				return "failed to create image";
			}
		} catch (Exception e) {
			return "Exception occurred. Cause : "+e.getLocalizedMessage();
		}
	}

	private BufferedImage bytesToBufferedImage(byte[] imageBytes) throws IOException {
		ByteArrayInputStream inputStream = new ByteArrayInputStream(imageBytes);
		return ImageIO.read(inputStream);
	}

	private String generateSvgFromExcel(Workbook workbook, Sheet sheet) {
		StringBuilder svgBuilder = new StringBuilder();

		// Loop through rows and cells to generate SVG content
		int maxColumnCount = 0;
		for (int i = 0; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			if (row != null) {
				int lastCellNum = row.getLastCellNum();
				if (lastCellNum > maxColumnCount) {
					maxColumnCount = lastCellNum;
				}
			}
		}
		System.out.println("maxColumnCount...." + maxColumnCount);
		int rowMaxPixels = 0;
		int colMaxPixels = 0;
		for (Row row : sheet) {
			float cellHeightInPoints = row.getHeightInPoints();
			int cellHeightInPixels = Math.round(cellHeightInPoints * 96 / 72);
			rowMaxPixels = rowMaxPixels + cellHeightInPixels;
		}
		for (int i = 0; i <= maxColumnCount; i++) {
			int cellWidth = Math.round(sheet.getColumnWidthInPixels(i));
			colMaxPixels = colMaxPixels + cellWidth;
		}
		// svg starts
		svgBuilder.append("<svg width=\"").append(colMaxPixels).append("\" height=\"").append(rowMaxPixels + 19) // change
																													// 19
				.append("\" xmlns=\"http://www.w3.org/2000/svg\">\n");

		double y = 0;
		List<CellRangeAddress> merged = sheet.getMergedRegions();
		Collections.sort(merged, new MergedRegionComparator());
		int mergedCount = sheet.getNumMergedRegions();
		List<Map<String, Object>> merge = new ArrayList<>();
		for (CellRangeAddress mergedCell : merged) {
			Map<String, Object> cells = new HashMap<>();
			Set<String> a1Notation = new HashSet<>();
//			int cellWidth = 0;
			int cellHeight = 0;
			int firstRow = mergedCell.getFirstRow();
			int firstCol = mergedCell.getFirstColumn();
			int lastRow = mergedCell.getLastRow();
			int lastCol = mergedCell.getLastColumn();
			for (int l = firstRow; l <= lastRow; l++) {
				int cellWidth = 0;
				for (int k = firstCol; k <= lastCol; k++) {
					cellWidth += Math.round(sheet.getColumnWidthInPixels(k));
					String cellRef = CellReference.convertNumToColString(k) + (l + 1);
					a1Notation.add(cellRef);
				}

				cellHeight += Math.round(sheet.getRow(l).getHeightInPoints() * 96 / 72);
//				a1Notation.add(null);
				cells.put("width", cellWidth);
			}

			cells.put("rowSpan", lastRow - firstRow);
			cells.put("colSpan", lastCol - firstCol);
			cells.put("height", cellHeight);
			cells.put("a1Notation", a1Notation);
			merge.add(cells);
		}
		for (int i = 0; i < mergedCount; i++) {
			System.out.println(merge.get(i));
		}
		for (Row row : sheet) {
			double cellWidth = 0;
			double x = 0;
			float cellHeightInPoints = row.getHeightInPoints();
			double cellHeight = Math.round(cellHeightInPoints * 96 / 72);
			y = y + cellHeight;
			for (int i = 0; i < maxColumnCount; i++) {
				Cell cell = row.getCell(i);
				x = x + cellWidth;
				cellWidth = Math.round(sheet.getColumnWidthInPixels(i));

				if (cell != null) {
					Object cellValue = getCellValueAsString(cell);
//					 System.out.println(cellRef);
					// Calculate cell position based on row and column indexes
					CellStyle cellStyle = cell.getCellStyle();
					Color backgroundColor = cellStyle.getFillForegroundColorColor();
					String hexColor = "#E0E0E0";
					if (backgroundColor instanceof XSSFColor) {
						XSSFColor xssfColor = (XSSFColor) backgroundColor;
						byte[] rgb = xssfColor.getRGB();
						hexColor = String.format("#%02X%02X%02X", rgb[0], rgb[1], rgb[2]);
						System.out.println(hexColor);
					}
					Font font = workbook.getFontAt(cellStyle.getFontIndex());
					String fontFamily = font.getFontName();
					int fontSize = Math.round(font.getFontHeightInPoints() * 96 / 72);
//					To get Font Color
					XSSFColor color = ((XSSFFont) font).getXSSFColor();
					String fontColor = "#000000";
					if (color != null) {
						byte[] rgb = color.getRGB();
						fontColor = String.format("#%02x%02x%02x", (rgb[0] < 0) ? (rgb[0] + 256) : (rgb[0]),
								(rgb[1] < 0) ? (rgb[1] + 256) : (rgb[1]), (rgb[2] < 0) ? (rgb[2] + 256) : (rgb[2]));
					}
					// text position
					double textX = x + 5;
					double textY = y + 15;
					/*
					 * if(cellStyle.getVerticalAlignment().toString().equals("TOP")) { textX = x+5;
					 * textY = y+fontSize; }else
					 * if(cellStyle.getVerticalAlignment().toString().equals("MIDDLE")) { textX =
					 * x+(cellWidth/2); textY = y+cellHeight/2; }else
					 * if(cellStyle.getVerticalAlignment().toString().equals("BOTTOM")){ textX =
					 * x+5; textY = y+cellHeight/1.5; }
					 * if(cellStyle.getAlignment().toString().equals("CENTER")) { textX =
					 * x+cellWidth/2 - cellWidth/8;
					 * 
					 * }else if(cellStyle.getAlignment().toString().equals("RIGHT")) { textX = x+15;
					 * }
					 */
					String textAnchor = mapHorizontalAlignmentToSVG(cellStyle.getAlignment());
					String dominantBaseLine = mapVerticalAlignmentToSVG(cellStyle.getVerticalAlignment());
					// Create SVG rect element for cell background
					svgBuilder.append("<rect x=\"").append(x).append("\" y=\"").append(y).append("\" width=\"")
							.append(cellWidth).append("\" height=\"").append(cellHeight).append("\" fill=\"")
							.append(hexColor).append("\" />\n");

					// Create SVG text element for cell value
					svgBuilder.append("<text x=\"").append(textX).append("\" y=\"").append(textY)
							.append("\" font-family=\"").append(fontFamily).append("\" font-size=\"").append(fontSize);
//							.append("\" dominant-baseline=\"").append(dominantBaseLine).append("\" text-anchor=\"")
//							.append(textAnchor);
					if (font.getBold()) {
						svgBuilder.append("\" font-weight=\"bold");
					}
					if (font.getItalic()) {
						svgBuilder.append("\" font-style=\"italic");
					}
					svgBuilder.append("\" fill=\"").append(fontColor).append("\">");
					if (!cellValue.equals("_")) {
						svgBuilder.append(cellValue);
					}
					svgBuilder.append("</text> \n");

				} else {
					svgBuilder.append("<rect x=\"").append(x).append("\" y=\"").append(y).append("\" width=\"")
							.append(cellWidth).append("\" height=\"").append(cellHeight).append("\" fill=\"")
							.append("#E0E0E0").append("\" />\n");
				}
			}
		}

		svgBuilder.append("</svg>");
		System.out.println(svgBuilder);
		return svgBuilder.toString();
	}

	private Object getCellValueAsString(Cell cell) {
		DataFormatter formatter = new DataFormatter();
		if (cell.getCellType() == CellType.STRING) {
			return cell.getStringCellValue();
		} else if (cell.getCellType() == CellType.NUMERIC) {
			if (DateUtil.isCellDateFormatted(cell)) {
				String format = formatter.formatCellValue(cell);
				return format;
			} else {
				return String.valueOf(cell.getNumericCellValue());
			}
		} else {
			return "";
		}
	}

	private byte[] convertSvgToImage(String svgContent) throws IOException, TranscoderException {
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

		// Create a PNG transcoder
		Transcoder transcoder = new PNGTranscoder();
		TranscoderInput input = new TranscoderInput(new java.io.StringReader(svgContent));
		TranscoderOutput output = new TranscoderOutput(outputStream);

		// Perform the conversion
		transcoder.transcode(input, output);

		// Close the stream and return the image bytes
		outputStream.close();
		return outputStream.toByteArray();
//        return outputStream.toByteArray();
	}

	private String mapHorizontalAlignmentToSVG(HorizontalAlignment alignment) {
		switch (alignment) {
		case LEFT:
			return "start";
		case CENTER:
			return "middle";
		case RIGHT:
			return "end";
		default:
			return "start";
		}
	}

	private String mapVerticalAlignmentToSVG(VerticalAlignment alignment) {
		switch (alignment) {
		case TOP:
			return "text-before-edge";
		case CENTER:
			return "middle";
		case BOTTOM:
			return "text-after-edge";
		default:
			return "text-before-edge";
		}
	}

}

class MergedRegionComparator implements Comparator<CellRangeAddress> {
	@Override
	public int compare(CellRangeAddress region1, CellRangeAddress region2) {
		return Integer.compare(region1.getFirstRow(), region2.getFirstRow());
	}
}
