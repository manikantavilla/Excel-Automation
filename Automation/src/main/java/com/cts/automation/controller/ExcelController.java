package com.cts.automation.controller;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

import com.cts.automation.model.CvsData;
import com.cts.automation.model.User;
import com.cts.automation.model.VendorData;
import com.cts.automation.service.ExcelService;
import com.fasterxml.jackson.databind.ObjectMapper;

import lombok.extern.slf4j.Slf4j;

@CrossOrigin(origins = "*", allowedHeaders = "*")
@RestController
@Slf4j
@RequestMapping("/api/v1.0")
public class ExcelController {


	@Autowired
	private ExcelService excelService;
	
	@Autowired
    private VendorData vendorData;
	
	@Autowired
    private CvsData cvsData;
	
	@GetMapping("/")
	public ModelAndView homePage() {
		return new ModelAndView("sample_forecast", "", null);
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
    
    @PostMapping("/createExcelFile")
    public ResponseEntity<byte[]> createExcelFile(@RequestPart("forecast_file") MultipartFile file, @RequestPart("filters") String filters) throws Exception {
    	User user = new ObjectMapper().readValue(filters, User.class);
    return excelService.insertDataIntoExcel(user);
    }
    
    @GetMapping("/vendors")
    public Map<String, Map<String, String>> getAllVendors() {   	    	
    	return vendorData.getVendor();
    }
    
    @GetMapping("/cvs")
    public Map<String, Map<String, String>> getAllCvsData() {   	    	
    	return cvsData.getCvs();
    }
    
    @GetMapping("/vendorId")
    public List<String> getVendorNameById() { 
    	List<String> result = new ArrayList<>();
    	
    	Map<String, Map<String, String>> ven= new HashMap<>();
    	ven = vendorData.getVendor();
    	for(Map.Entry<String, Map<String, String>> entry : ven.entrySet()) {
    		for(Map.Entry<String, String> entry2 : entry.getValue().entrySet()) {
    			if(entry.getKey().equals("208876")) {
    				result.add(entry2.getValue());
    			}
    		}
    	}
    	return result;
    }

  
}

