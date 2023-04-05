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

import com.cts.automation.model.Amendment;
import com.cts.automation.model.CostCenters;
import com.cts.automation.model.CvsData;
import com.cts.automation.model.SOWPath;
import com.cts.automation.model.SheetName;
import com.cts.automation.model.SubmitPath;
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
	
	@Autowired
	private SOWPath SowPath;
	
	@Autowired
	private SubmitPath submiTPath;
	
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
    	return excelService.FileNames();
    }
    
    
    @GetMapping("/getSowPath")
    public String getsowPath() {
    	return SowPath.getSowPath();
    }
    
    @GetMapping("/getSubmitPath")
    public String getsubmitPath() {
    	return submiTPath.getSubmitPath();
    }
    
 
    
    
    @PostMapping("/getTempExcelData")
    public List<Map<String, Object>> readAmendmentData(@RequestPart("forecast_file") MultipartFile file,@RequestPart("filters") Amendment user) throws Exception {
        return excelService.ReadAmendmentData(file,user);
    }
    
    @PostMapping("/createAmendmentWordFile")
    public ResponseEntity<byte[]> createAmendmentWordFile(@RequestPart("forecast_file") MultipartFile file,@RequestPart("filters") String filters) throws Exception {
    	Amendment user = new ObjectMapper().readValue(filters, Amendment.class);
        return excelService.insertDataIntoAmendmentWord(file,user);
    }
    
    @PostMapping("/createAmendmentExcelFile")
    public ResponseEntity<byte[]> createAmendmentExcelFile(@RequestPart("forecast_file") MultipartFile file, @RequestPart("filters") String filters) throws Exception {
    	Amendment user = new ObjectMapper().readValue(filters, Amendment.class);
    return excelService.insertDataIntoAmendmentExcel(user);
    }
}

