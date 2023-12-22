package com.excel.image.controllers;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.excel.image.RequestBody.JsonToExcelConvertRequest;
import com.excel.image.services.ExcelService;

@RestController
@RequestMapping("")
public class ExcelController {
	@Autowired
	private ExcelService excelService;

	@PostMapping("/json/excel/convert")
	public String jsonToExcelConvertor(@RequestBody JsonToExcelConvertRequest jsonToExcelConvertRequest) {
		return excelService.jsonToExcelConvertor(jsonToExcelConvertRequest);
	}
	
	@PostMapping("/excel/image/convert")
	public String excelToImageConvertor(@RequestBody JsonToExcelConvertRequest jsonToExcelConvertRequest) {
		return excelService.excelToImageConvertor(jsonToExcelConvertRequest);
	}
}
