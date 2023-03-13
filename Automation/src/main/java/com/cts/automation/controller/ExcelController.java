package com.cts.automation.controller;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

import com.cts.automation.model.CostCenters;
import com.cts.automation.model.CvsData;
import com.cts.automation.model.SheetName;
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
	
	@Value("${costCenters}")    
	String[] costCenters;
	
	@Value("${sheetNames}")    
	String[] sheetName;


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
        log.info("Word");
    	User user = new ObjectMapper().readValue(filters, User.class);
        return excelService.insertDataIntoWord(file,user);
    }
    
    @PostMapping("/createExcelFile")
    public ResponseEntity<byte[]> createExcelFile(@RequestPart("forecast_file") MultipartFile file, @RequestPart("filters") String filters) throws Exception {
    	log.info("excel");
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
    
    @GetMapping("/cost-centers")
    public List<CostCenters> getCostCenters() {
        List<CostCenters> result = new ArrayList<>();
        for (String code : costCenters) {
            result.add(new CostCenters(code));
        }
        return result;
    } 
    
    @GetMapping("/sheetNames")
    public List<SheetName> getSheetNames() {
        List<SheetName> result = new ArrayList<>();
        for (String sheetN : sheetName) {
            result.add(new SheetName(sheetN));
        }
        return result;
    }
    
    @GetMapping("/sowName")
    public List getSOWName() {
    	log.info("file");
    	String NewSFName = new String();
    	
    	String NeWSFName = new String();
    	
    	 String NewSowName = new String();
    	
    	for (int i = 0; i < ExcelService.defaultName.length(); i++) {
    		NewSowName += ExcelService.defaultName.charAt(i);
  
            if (i == 3) {
            	NewSowName += ExcelService.sowName;
            }
        }
    	NewSowName += ".docx";
        log.info(NewSowName);
        
        
			  
        for (int i = 0; i < ExcelService.defaultName.length(); i++) {
        	NewSFName += ExcelService.defaultName.charAt(i);
  
            if (i == 2) {
            	NewSFName += "SF";
            }
        }
        
        for (int i = 0; i < NewSFName.length(); i++) {
        	NeWSFName += NewSFName.charAt(i);
  
        	if (i == 5) {
        		NeWSFName +=ExcelService.sowName;
            }
        }
        NeWSFName += ".xlsx";
        log.info(NeWSFName);
    	
        List fileNames = new ArrayList<>();
        fileNames.add(NewSowName);
        fileNames.add(NeWSFName);
        NewSowName = "";
        NeWSFName = "";
        NewSFName = "";
        return fileNames;
    }
    
}

