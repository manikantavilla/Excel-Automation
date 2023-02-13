package com.cts.automation.controller;

import java.util.List;
import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.support.ResourceBundleMessageSource;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

import com.cts.automation.model.User;
import com.cts.automation.service.ExcelService;
import com.fasterxml.jackson.databind.ObjectMapper;

@CrossOrigin(origins = "*", allowedHeaders = "*")
@RestController
@RequestMapping("/api/v1.0")
public class ExcelController {

	@Autowired
	private ExcelService excelService;
	
	@GetMapping("/")
	public ModelAndView homePage() {
		return new ModelAndView("forecast", "", null);
	}
	
	
	@PostMapping("/getDataFromExcel")
    public List<Map<String, Object>> readDataFromExcel(@RequestPart("forecast_file") MultipartFile file,@RequestPart("filters") User user) throws Exception {
        return excelService.ReadBasedOnCondition(file,user);
    }

	@PostMapping("/createWordFile")
    public ResponseEntity<byte[]> createWordFile(@RequestPart("forecast_file") MultipartFile file,@RequestPart("filters") String filters) throws Exception {
        User user = new ObjectMapper().readValue(filters, User.class);
        return excelService.insertDataIntoWord(file,user);
    }


}

