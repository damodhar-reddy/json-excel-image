package com.excel.image.RequestBody;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class JsonToExcelConvertRequest {
	
	private String sourcePath;

	private String destination;
	
	private String sheetName;
}
