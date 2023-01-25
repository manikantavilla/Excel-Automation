package com.cts.automation.service;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.cts.automation.model.User;

import lombok.extern.slf4j.Slf4j;


@Service
@Slf4j
public class ExcelService {

	public List<Map<String, Object>> ReadBasedOnCondition(MultipartFile file) throws Exception {
//		FileInputStream inputStream = (FileInputStream) file.getInputStream();;
//		Workbook workbook = new XSSFWorkbook(inputStream);
		Workbook workbook = new XSSFWorkbook(file.getInputStream());
		Sheet sheet = workbook.getSheetAt(0);
		String columnName = "Course";
		int columnIndex = -1;
		Row headerRow = sheet.getRow(0);
		for(Cell cell : headerRow) {
			if(cell.getCellType() == CellType.STRING) {
		    if(cell.getStringCellValue().equalsIgnoreCase(columnName)) {
		        columnIndex = cell.getColumnIndex();
		        break;
		    }
			}
			else if(cell.getCellType() == CellType.NUMERIC) {
		    if(Double.toString(cell.getNumericCellValue()).equalsIgnoreCase(columnName))
		    {
		        columnIndex = cell.getColumnIndex();
		        break;
		    }
			}
		}
		


		List<Row> rows = new ArrayList<Row>();	
//		for(Row row : sheet) {
		for(int i= 0; i< sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
		    Cell cell = row.getCell(columnIndex);
		    if(cell != null) {
		    if(cell.getStringCellValue().equalsIgnoreCase("CSE")) 
//		    if(Double.toString(cell.getNumericCellValue()).equalsIgnoreCase("N66YYYY"))	
		    {
		        rows.add(row);
		    }
		    }
		    else {
		    	break;
		    }
		}

		List<Map<String, Object>> rowsData = new ArrayList<Map<String, Object>>();
		for(Row row : rows) {
		    Map<String, Object> rowData = new HashMap<String, Object>();
		    for (Cell cell : row) {
		        int columnIndexx = cell.getColumnIndex();
		        String columnNamee = headerRow.getCell(columnIndexx).getStringCellValue();
		        if (cell.getCellType() == CellType.NUMERIC) {
			        if (DateUtil.isCellDateFormatted(cell)) {
			            rowData.put(columnNamee, cell.getDateCellValue().toString());
			        } else {
			            rowData.put(columnNamee, Double.toString(cell.getNumericCellValue()));
			        }
			    } else if (cell.getCellType() == CellType.STRING) {
			        rowData.put(columnNamee, cell.getStringCellValue());
			    }
			    else if (cell.getCellType() == CellType.FORMULA) {
			        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
			        CellValue cellValue = evaluator.evaluate(cell);
			        switch (cellValue.getCellType()) {
			            case NUMERIC:
//			                System.out.println(cellValue.getNumberValue());
			            	rowData.put(columnNamee, Double.toString(cellValue.getNumberValue()));
			                break;
			            case STRING:
//			                System.out.println(cellValue.getStringValue());
			                rowData.put(columnNamee, cellValue.getStringValue());
			                break;
			    }
		    }
		    }
		    rowsData.add(rowData);
		}
		return rowsData;
	}

	

//    public ResponseEntity<byte[]> insertDataIntoWord() throws Exception {
//        XWPFDocument doc = new XWPFDocument();
//        XWPFParagraph p = doc.createParagraph();
//        for(Map<String, Object> row : ReadBasedOnCondition()) {
//            if(row.get("Course").equals("CSE")) {
//                p.createRun().setText("Name : " + row.get("Name") + " ");
//                p.createRun().setText("Employee ID : " + row.get("Student ID") + " ");
//                p = doc.createParagraph();
//            }
//        }
//        ByteArrayOutputStream out = new ByteArrayOutputStream();
//        doc.write(out);
//        out.close();
//
//        HttpHeaders headers = new HttpHeaders();
//        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
//        headers.setContentDispositionFormData("attachment", "SOW.docx");
//        headers.setContentLength(out.toByteArray().length);
//
//        return new ResponseEntity<byte[]>(out.toByteArray(), headers, HttpStatus.OK);
//    }

	public ResponseEntity<byte[]> insertDataIntoWord(MultipartFile file,User user) throws Exception {
	    FileInputStream inputStream = new FileInputStream(new File("C:\\Users\\2066253\\eclipse-workspace\\COG2023-0XX_CCCC86_SOW_Business for Active Health_2023_Copy.docx"));
	    XWPFDocument doc = new XWPFDocument(inputStream);
	    List<XWPFParagraph> paragraphs = doc.getParagraphs();
	    for (XWPFParagraph paragraph : paragraphs) {
	        List<XWPFRun> runs = paragraph.getRuns();
	        for (XWPFRun run : runs) {
	            String text = run.getText(0);
	            for(Map<String, Object> row : ReadBasedOnCondition(file)) {
	                if(text != null && text.contains("<Header>")) {
//	                    text = text.replace("COG2023-0XX ", (String) row.get("Course"));
	                	text = text.replace("<Header>", (String) row.get("Name"));
	                    log.info("Okay Change Success");
	                    run.setText(text, 0);
	                }
	                if(text != null && text.contains("DATE")) {
	                    text = text.replace("DATE", user.getStartdate());
	                    log.info(user.getStartdate());
	                    run.setText(text, 0);
	                }
	                if(text != null && text.contains("<YEAR>")) {
	                    text = text.replace("<YEAR>", (String) user.getYear());
	                    run.setText(text, 0);
	                }

	                if(text != null && text.contains("END")) {
	                    text = text.replace("END", user.getEnddate());
	                    log.info(user.getEnddate());
	                    run.setText(text, 0);
	                }
	                if(text != null && text.contains("TOTALAMOUNT")) {
	                    text = text.replace("TOTALAMOUNT", Double.toString(user.getTotalamount()));
	                    log.info(Double.toString(user.getTotalamount()));
	                    run.setText(text, 0);
	                }
	            }
	        }
	    }
	    
	    
	    
//	    for iterating all tables
	        
	    List<XWPFTable> tables = doc.getTables();
	    for (XWPFTable table : tables) {
	        for (XWPFTableRow row : table.getRows()) {
	            for (XWPFTableCell cell : row.getTableCells()) {
	                if (cell.getText().contains("<VendorProjectTeamMember1>")) {
	                    cell.removeParagraph(0);
	                    XWPFParagraph newPara = cell.addParagraph();
	                    XWPFRun run = newPara.createRun();
	                    run.setText(user.getVendorprojectteammember1());
	                }
	                if (cell.getText().contains("<VendorProjectTeamMember2>")) {
	                    cell.removeParagraph(0);
	                    XWPFParagraph newPara = cell.addParagraph();
	                    XWPFRun run = newPara.createRun();
	                    run.setText(user.getVendorprojectteammember2());
	                }
	                if (cell.getText().contains("<VendorProjectTeamMember3>")) {
	                    cell.removeParagraph(0);
	                    XWPFParagraph newPara = cell.addParagraph();
	                    XWPFRun run = newPara.createRun();
	                    run.setText(user.getVendorprojectteammember3());
	                }
	                if (cell.getText().contains("<CVSProjectTeamMember1>")) {
	                    cell.removeParagraph(0);
	                    XWPFParagraph newPara = cell.addParagraph();
	                    XWPFRun run = newPara.createRun();
	                    run.setText(user.getCvsprojectteammember1());
	                }
	                if (cell.getText().contains("TOTALAMOUNT")) {
	                    cell.removeParagraph(0);
	                    XWPFParagraph newPara = cell.addParagraph();
	                    XWPFRun run = newPara.createRun();
	                    run.setText(Double.toString(user.getTotalamount()));
	                    log.info(Double.toString(user.getTotalamount()));
	                }
	            }
	        }
	    }
	    
//	    for iterating one table
	    
	    
//		 XWPFTable table = doc.getTables().get(0);
//		 for (int i = 0; i < table.getNumberOfRows(); i++) {
//		     XWPFTableRow row = table.getRow(i);
//		     for (int j = 0; j < row.getTableCells().size(); j++) {
//		         XWPFTableCell cell = row.getCell(j);
//		         if (cell.getText().contains("Service Fee")) {
//		             cell.removeParagraph(0);
//		             cell.addParagraph().createRun().setText("12345");
//		         }
//		     }
//		 }
	    
	    ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
	    doc.write(byteArrayOutputStream);
	    HttpHeaders headers = new HttpHeaders();
	    headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
	    headers.add("Content-Disposition", "attachment; filename=SOW_Document.docx");
	    headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");
	    ResponseEntity<byte[]> response = new ResponseEntity<byte[]>(byteArrayOutputStream.toByteArray(), headers, HttpStatus.OK);
	    byteArrayOutputStream.close();
	    return response;
	}

}
