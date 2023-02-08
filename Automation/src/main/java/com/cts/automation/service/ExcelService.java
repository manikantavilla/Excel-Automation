package com.cts.automation.service;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.lang.reflect.InvocationTargetException;
import java.text.DateFormatSymbols;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.temporal.TemporalAccessor;
import java.time.temporal.TemporalAdjusters;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
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

	public List<Map<String, Object>> ReadBasedOnCondition(MultipartFile file, User user) throws Exception {
//		FileInputStream inputStream = (FileInputStream) file.getInputStream();;
//		Workbook workbook = new XSSFWorkbook(inputStream);
		Workbook workbook = new XSSFWorkbook(file.getInputStream());
		Sheet sheet = workbook.getSheetAt(1);
		String columnName = "Cost Center";
//		String startDateColumnName = "Start Date";
//		int startDateColumnIndex = -1;
		int columnIndex = -1;
		Row headerRow = sheet.getRow(10);
		for(Cell cell : headerRow) {
			if(cell.getCellType() == CellType.STRING) {
;		    if(cell.getStringCellValue().equalsIgnoreCase(columnName)) {
		        columnIndex = cell.getColumnIndex();
//		        break;
		    }
//			if (cell.getStringCellValue().equalsIgnoreCase(startDateColumnName)) {
//				startDateColumnIndex = cell.getColumnIndex();
////				break;
//			}
			}
			else if(cell.getCellType() == CellType.NUMERIC) {
		    if(Double.toString(cell.getNumericCellValue()).equalsIgnoreCase(columnName))
		    {
		        columnIndex = cell.getColumnIndex();
//		        break;
		    }
//		    if(Double.toString(cell.getNumericCellValue()).equalsIgnoreCase(startDateColumnName))
//		    {
//		    	startDateColumnIndex = cell.getColumnIndex();
////		        break;
//		    }
			}
		}
		


		List<Row> rows = new ArrayList<Row>();	
//		for(Row row : sheet) {
		for(int i= 11; i< sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
		    Cell cell = row.getCell(columnIndex);
		    if(cell != null) {
		    if(cell.getStringCellValue().equalsIgnoreCase(user.getCostcenter())) 
		    {
		        rows.add(row);
		    }
		    }
		    else {
		    	break;
		    }
		}
		
		
		
		List<Map<String, Object>> rowsData = new ArrayList<Map<String, Object>>();
		for (Row row : rows) {
		    Map<String, Object> rowData = new HashMap<String, Object>();
		    for (Cell cell : row) {
		        int columnIndexx = cell.getColumnIndex();
//		        String columnNamee = headerRow.getCell(columnIndexx).getStringCellValue();
		        Cell headerCell = headerRow.getCell(columnIndexx);
		        String columnNamee = "";
		        switch (headerCell.getCellType()) {
		            case NUMERIC:
		                if (DateUtil.isCellDateFormatted(headerCell)) {
		                    columnNamee = headerCell.getDateCellValue().toString();
		                } else {
		                    columnNamee = String.valueOf(headerCell.getNumericCellValue());
		                }
		                break;
		            case STRING:
		                columnNamee = headerCell.getStringCellValue();
		                break;
		            case BOOLEAN:
		                columnNamee = String.valueOf(headerCell.getBooleanCellValue());
		                break;
		        }
		        switch (cell.getCellType()) {
		            case NUMERIC:
		                if (DateUtil.isCellDateFormatted(cell)) {
		                    rowData.put(columnNamee, cell.getDateCellValue().toString());
		                } else {
		                    rowData.put(columnNamee, cell.getNumericCellValue());
		                }
		                break;
		            case STRING:
		                rowData.put(columnNamee, cell.getStringCellValue());
		                break;
		            case BOOLEAN:
		                rowData.put(columnNamee, cell.getBooleanCellValue());
		                break;
		            case FORMULA:
		                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		                CellValue cellValue = evaluator.evaluate(cell);
		                if (cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)")) {
		                    rowData.put(columnNamee, "$" + cellValue.getNumberValue());
		                } else {
//		                    switch (cellValue.getCellType()) {
		                	switch (evaluator.evaluateFormulaCell(cell)) {
		                        case NUMERIC:
//		                            rowData.put(columnNamee, cellValue.getNumberValue());
		                        	rowData.put(columnNamee, cell.getNumericCellValue());
		                            break;
		                        case STRING:
		                            rowData.put(columnNamee, cell.getStringCellValue());
		                            break;
		                        case BOOLEAN:
		                            rowData.put(columnNamee, cellValue.getBooleanValue());
		                            break;
		                    }
		                }
		                break;
		        }
		    }
		    rowsData.add(rowData);
		}


		return rowsData;
	}
	

	public ResponseEntity<byte[]> insertDataIntoWord(MultipartFile file,User user) throws Exception {
	    FileInputStream inputStream = new FileInputStream(new File("C:\\Users\\2066253\\eclipse-workspace\\COG2023-0XX_CCCC86_SOW_Business for Active Health_2023_20230203.docx"));
	    XWPFDocument doc = new XWPFDocument(inputStream);
	    List<Map<String, Object>> rowsData = new ArrayList<Map<String, Object>>();
	    rowsData = ReadBasedOnCondition(file,user);	    
	    List<XWPFParagraph> paragraphs = doc.getParagraphs();
	    for (XWPFParagraph paragraph : paragraphs) {
	        List<XWPFRun> runs = paragraph.getRuns();
	        for (XWPFRun run : runs) {
	            String text = run.getText(0);
	            for(Map<String, Object> row : rowsData) {
	            	log.info(text);
	                if(text != null && text.contains("<contract_id>")) {
	                	text = text.replace("<contract_id>", (String) row.get("Contract#"));
	                    run.setText(text, 0);
	                }
	                if(text != null && text.contains("\\<sow_effective_date\\>")) {
	                	
	                	String dateString = (String) row.get("Start Date");
	                	log.info(dateString);
	                	System.out.println(row.get("Start Date"));
	                	SimpleDateFormat inputFormat = new SimpleDateFormat("d-MMM-yy");
	                	SimpleDateFormat outputFormat = new SimpleDateFormat("MMMM dd, yyyy");
	                	Date date = inputFormat.parse(dateString);
	                	String outputStartDateString = outputFormat.format(date);
	                	
	                    text = text.replace("<sow_effective_date>", outputStartDateString);
	                    run.setText(text, 0);
	                }
	                if(text != null && text.contains("<current_year>.")) {
	                	
	                	String dateString =  (String) row.get("End Date");
	                	SimpleDateFormat inputFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss z yyyy");
	                	SimpleDateFormat outputFormat = new SimpleDateFormat("yyyy");
	                	Date date = inputFormat.parse(dateString);
	                	String outputYearString = outputFormat.format(date);
	                	
	                    text = text.replace("<current_year>.", outputYearString);
	                    run.setText(text, 0);
	                }

	                if(text != null && text.contains("END")) {
	                	
	                	String dateString =  (String) row.get("End Date");
	                	SimpleDateFormat inputFormat = new SimpleDateFormat("dd-MMM-yy");
	                	SimpleDateFormat outputFormat = new SimpleDateFormat("MMMM dd, yyyy");
	                	Date date = inputFormat.parse(dateString);
	                	String outputEndDateString = outputFormat.format(date);
	                	
	                    text = text.replace("END",outputEndDateString);
	                    run.setText(text, 0);
	                }
	                if(text != null && text.contains("<budget_amount>")) {
	                	text = text.replace("<budget_amount>", "$"+String.format("%.2f", (double) row.get("Totals")));
	                    run.setText(text, 0);
	                }
	                break;
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
//	                if (cell.getText().contains("TOTALAMOUNT")) {
//	                    cell.removeParagraph(0);
//	                    XWPFParagraph newPara = cell.addParagraph();
//	                    XWPFRun run = newPara.createRun();
//	                    run.setText((String) row.get("Totals"));
//	                    log.info(Double.toString(user.getTotalamount()));
//	                }
	            }
	        }
	    }
	    
	    
	  //Create a new table
	    List<String> months = Arrays.asList("January", "February", "March","April","May","June","July","August","September","October","November","December");
	    XWPFTable newTable = doc.createTable();
	    XWPFTableRow row1 = newTable.createRow();
        XWPFTableCell cell0 = row1.createCell();
        XWPFTableCell cell01 = row1.createCell();
        XWPFParagraph para0 = cell0.getParagraphs().get(0);
        XWPFRun run0 = para0.createRun();
        run0.setBold(true);
        run0.setText("Deliverable");
        
        XWPFParagraph para01 = cell01.getParagraphs().get(0);
        XWPFRun run01 = para01.createRun();
        run01.setBold(true);
        run01.setText("Date to be Complete");

	    int tableIndex = -1;
	    for (int i = 0; i < tables.size(); i++) {
	        XWPFTable table = tables.get(i);
	        if (table.getText().contains("Deliverable")) {
	            tableIndex = i;
	            log.info(Double.toString(tableIndex));
	            break;
	        }
	    }
	    //Remove the old table
	    if (tableIndex != -1) {
	    	doc.removeBodyElement(tableIndex-1);
	   
	    	Date StartDate = user.getStartdate();
	    	SimpleDateFormat StartDateFormat = new SimpleDateFormat("MMMM");
	    	String startDate = StartDateFormat.format(StartDate);
	    	
	    	Date EndDate = user.getEnddate();
	    	SimpleDateFormat EndDateFormat = new SimpleDateFormat("MMMM");
	    	String endDate = EndDateFormat.format(EndDate);

	    	int monthStartIndex = 0;
	    	monthStartIndex = months.indexOf(startDate);
	    	int monthEndIndex = 0;
	    	if(endDate.length() > 0) {
	    		monthEndIndex = months.indexOf(endDate);
				for (int i = monthStartIndex+1; i <= monthEndIndex + 1; i++) {
					XWPFTableRow row = newTable.createRow();
					XWPFTableCell cell1 = row.createCell();
					LocalDate date = LocalDate.of(LocalDate.now().getYear(), i, 1);
					cell1.setText("Services for the Month of " + months.get(i - 1) + " "
							+ date.with(TemporalAdjusters.lastDayOfMonth()).getYear());
					XWPFTableCell cell2 = row.createCell();
					cell2.setText(date.with(TemporalAdjusters.lastDayOfMonth()).getDayOfMonth() + "-"
							+ months.get(i - 1).substring(0, 3) + "-"
							+ date.with(TemporalAdjusters.lastDayOfMonth()).getYear());
				}
			}
			else {
				XWPFTableRow row = newTable.createRow();
				XWPFTableCell cell1 = row.createCell();
				LocalDate date = LocalDate.of(LocalDate.now().getYear(), monthStartIndex, 1);
				cell1.setText("Services for the Month of " + months.get(monthStartIndex) + " "
						+ date.with(TemporalAdjusters.lastDayOfMonth()).getYear());
				XWPFTableCell cell2 = row.createCell();
				cell2.setText(date.with(TemporalAdjusters.lastDayOfMonth()).getDayOfMonth() + "-"
						+ months.get(monthStartIndex - 1).substring(0, 3) + "-"
						+ date.with(TemporalAdjusters.lastDayOfMonth()).getYear());
			}
		    
	        //Insert the new table at the same position as the old table
		    doc.setTable(tableIndex, newTable);
	    }
	    
	    List<XWPFTable> table = doc.getTables();
	    int lastTableIndex = table.size() - 1;
	    doc.removeBodyElement(lastTableIndex);

	    
	    ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
	    doc.write(byteArrayOutputStream);
	    HttpHeaders headers = new HttpHeaders();
	    headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
	    headers.add("Content-Disposition", "attachment; filename=SOW_Document1.docx");
	    headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");
	    ResponseEntity<byte[]> response = new ResponseEntity<byte[]>(byteArrayOutputStream.toByteArray(), headers, HttpStatus.OK);
	    byteArrayOutputStream.close();
	    return response;
	}

}
