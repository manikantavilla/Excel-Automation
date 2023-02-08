package com.cts.automation.controller;

import java.util.List;
import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.support.ResourceBundleMessageSource;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

import com.cts.automation.model.User;
import com.cts.automation.service.ExcelService;

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
    public List<Map<String, Object>> readDataFromExcel(@RequestPart("test_file") MultipartFile file,@RequestPart("test_json") User user) throws Exception {
        return excelService.ReadBasedOnCondition(file,user);
    }

    @PostMapping("/createWordFile")
    public ResponseEntity<byte[]> createWordFile(@RequestPart("test_file") MultipartFile file, @RequestPart("test_json") User user) throws Exception {
//        return excelService.insertDataIntoWord(file,user);
    
    //public ResponseEntity<byte[]> createWordFile(@RequestParam("user") User user,@RequestParam("excelFile") MultipartFile file) throws Exception {
    	return excelService.insertDataIntoWord(file,user);
    }
}

