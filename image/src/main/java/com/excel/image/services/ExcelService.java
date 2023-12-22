package com.excel.image.services;

import com.excel.image.RequestBody.JsonToExcelConvertRequest;

public interface ExcelService {

	public String jsonToExcelConvertor(JsonToExcelConvertRequest jsonToExcelConvertRequest);

	public String excelToImageConvertor(JsonToExcelConvertRequest jsonToExcelConvertRequest);
}
