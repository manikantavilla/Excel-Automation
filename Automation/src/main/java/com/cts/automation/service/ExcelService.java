package com.cts.automation.service;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.math.BigInteger;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.temporal.TemporalAdjusters;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVerticalJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.cts.automation.model.Amendment;
import com.cts.automation.model.AmendmentPath;
import com.cts.automation.model.CvsData;
import com.cts.automation.model.SOWPath;
import com.cts.automation.model.SubmitPath;
import com.cts.automation.model.User;
import com.cts.automation.model.VendorData;

import lombok.extern.slf4j.Slf4j;

@Service
@Slf4j
public class ExcelService {

	@Autowired
	private VendorData vendorData;

	@Autowired
	private CvsData cvsData;

	@Autowired
	private SOWPath SowPath;

	@Autowired
	private SubmitPath submiTPath;

	@Autowired
	private AmendmentPath amendmentPath;

	public static double budgetAmount = 0.0;

	public static double totalBudgetAmount = 0.0;

	public static Integer amendmentCount = 0;

	public static String sowName = "";

	public static String amendmentName = "";

	public static String defaultName = new String();

	public List<Map<String, Object>> ReadBasedOnCondition(MultipartFile file, User user) throws Exception {
		
		// Read the file as a Workbook
		
		Workbook workbook = new XSSFWorkbook(file.getInputStream());
		
		// Get the sheet based on the provided sheet name in the User object
		
		Sheet sheet = workbook.getSheet(user.getSheetName());
		
		// Check if the sheet exists
		
		if (sheet != null) {
		
			// Define the column to be searched for
			
			String columnName = "Cost Center";
			int columnIndex = -1;
			
			// Get the header row of the sheet
			
			Row headerRow = sheet.getRow(10);
			
			// Find the index of the column to be searched for
			
			for (Cell cell : headerRow) {
				if (cell.getCellType() == CellType.STRING) {
					if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
						columnIndex = cell.getColumnIndex();
					}
				} else if (cell.getCellType() == CellType.NUMERIC) {
					if (Double.toString(cell.getNumericCellValue()).equalsIgnoreCase(columnName)) {
						columnIndex = cell.getColumnIndex();
					}
				}
			}

			// Initialize an array to store matching rows
			
			List<Row> rows = new ArrayList<Row>();
			int dt = 0;
			
			// Loop through all the rows in the sheet
			
			for (int i = 11; i < sheet.getPhysicalNumberOfRows(); i++) {
			
				// Get the current row
				
				Row row = sheet.getRow(i);
				
				// Get the cell corresponding to the searched column
				
				Cell cell = row.getCell(columnIndex);
				if (cell != null) {
				
					// Get the cells corresponding to the start date and end date columns
					
					Cell cell1 = row.getCell(columnIndex + 1);
					Cell cell2 = row.getCell(columnIndex + 2);
					
					// Check if the start date is not null
					
					if (cell1.getDateCellValue() != null) {
					
						// Parse the start date and end date
						
						Date uDate = user.getStartDate();
						Date eDate = cell1.getDateCellValue();
						Date endDate = user.getEndDate();
						Date ExcelEndDate = cell2.getDateCellValue();
						SimpleDateFormat printFormat = new SimpleDateFormat("dd-MM-yyyy");
						SimpleDateFormat outputFormat = new SimpleDateFormat("yyyy");
						String outputInputDateString = outputFormat.format(uDate);
						String outputExcelDateString = outputFormat.format(eDate);
						String PrintDate = printFormat.format(uDate);
						String PrintEndDate = printFormat.format(eDate);
						
						// Check if the start date is before the end date
						
						if (uDate.before(endDate)) {
						
							// Check if the year of the start date matches the year of the current row
							
							if (outputExcelDateString.equals(outputInputDateString)) {
							
								// Check if the value in the searched column matches the provided cost center
								
								if (cell.getStringCellValue().equalsIgnoreCase(user.getCostCenter())) {
								
									// Add the matching row to the list of rows
									
									rows.add(row);
								}
							} else {
								
								// Set the dt variable to 1 to indicate that no matching date was found
								
								dt = 1;
								log.info("No Date Found!!! Given Date is " + PrintDate);
								break;
							}
						} else {
							
							// Set the dt variable to 1 to indicate that the start date is after the end date
							
							dt = 1;
							log.info("Start Date is greater than End Date!!! Given Start Date is " + PrintDate
									+ " Given End Date is " + PrintEndDate);
							break;
						}
					}
				} else {
					break;
				}
			}
			
			//Checking if the rows are empty or dt is empty
			
			if (rows.size() <= 0 && dt < 1) {
				log.info("No Data Found for this Cost Center!!! Given Cost Center is  " + user.getCostCenter());
			}

			// This method converts the data in the Excel sheet into a List of Maps.
			// Each row in the sheet is represented by a Map, with the key being the column header and the value being the cell data.
			// The List of Maps is returned at the end of the method.
			
			List<Map<String, Object>> rowsData = new ArrayList<Map<String, Object>>();
			
			// Iterate over each row in the Matched rows
			
			for (Row row : rows) {
				
				// Create a new Map to store the data for this row
				
				Map<String, Object> rowData = new HashMap<String, Object>();
				
				// Iterate over each cell in the row
				
				for (Cell cell : row) {
					
					// Get the index of the current cell's column
					
					int columnIndexx = cell.getColumnIndex();
					
					// Get the corresponding header cell for this column
					
					Cell headerCell = headerRow.getCell(columnIndexx);
					
					// Get the column name from the header cell
					
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
					
					// Get the data for this cell and add it to the Map for this row
					
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
						
						// Check if the cell format is a currency format
						
						if (cell.getCellStyle().getDataFormat() == HSSFDataFormat
								.getBuiltinFormat("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)")) {
							rowData.put(columnNamee, "$" + cellValue.getNumberValue());
						} else {
							
							// Get the data for the formula cell and add it to the Map for this row
							
							switch (evaluator.evaluateFormulaCell(cell)) {
							case NUMERIC:
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
				
				// Add the Map for this row to the List of Maps
				
				rowsData.add(rowData);
			}

			// Return the List of Maps representing the data in the sheet
			
			return rowsData;
		} else {
			
			// If the sheet was not found, return an empty List of Maps
			
			List<Map<String, Object>> rowsData = new ArrayList<Map<String, Object>>();
			log.info("Sheet Not Found!!! Given Sheet Name is " + user.getSheetName());
			return rowsData;
		}
	}

	public ResponseEntity<byte[]> insertDataIntoWord(MultipartFile file, User user) throws Exception {
		FileInputStream inputStream = new FileInputStream(new File(SowPath.getSowPath()));
		XWPFDocument doc = new XWPFDocument(inputStream);
		List<Map<String, Object>> rowsData = new ArrayList<Map<String, Object>>();
		rowsData = ReadBasedOnCondition(file, user);
		if (rowsData.size() > 0) {

			List<XWPFTable> tables = doc.getTables();

			List<String> vendorTeamList = user.getVendorTeam();
			List<String> vendorNameList = new ArrayList<String>();
			List<String> vendorRoleList = new ArrayList<String>();

			Map<String, Map<String, String>> vendorMap = vendorData.getVendor();
			for (String vendorId : vendorTeamList) {
				Map<String, String> vendorDetails = vendorMap.get(vendorId);
				String vendorName = vendorDetails.get("name");
				String vendorRole = vendorDetails.get("role");
				vendorNameList.add(vendorName);
				vendorRoleList.add(vendorRole);

			}

			List<String> CVSTeamList = user.getCvsTeam();
			List<String> CVSNameList = new ArrayList<String>();
			List<String> CVSRoleList = new ArrayList<String>();

			Map<String, Map<String, String>> cvsMap = cvsData.getCvs();
			for (String cvsId : CVSTeamList) {
				Map<String, String> cvsDetails = cvsMap.get(cvsId);
				String cvsName = cvsDetails.get("name");
				String cvsRole = cvsDetails.get("role");
				CVSNameList.add(cvsName);
				CVSRoleList.add(cvsRole);

			}

			int rows = Math.max(vendorNameList.size(), CVSNameList.size());

			XWPFTable nTable = doc.createTable(rows + 1, 2);

			int size = 9800;
			nTable.setWidth("100%");
			nTable.getCTTbl().addNewTblPr().addNewTblW().setW(BigInteger.valueOf(size));

			XWPFTableRow nTableheaderRow = nTable.getRow(0);
			nTableheaderRow.getCell(0).setColor("CCE1FD");
			nTableheaderRow.getCell(1).setColor("CCE1FD");
			XWPFParagraph para = nTableheaderRow.getCell(0).getParagraphs().get(0);
			XWPFRun nTableRun = para.createRun();
			nTableRun.setBold(true);
			nTableRun.setText("Vendor Project Team:");
			nTableRun.setFontFamily("Arial");
			nTableRun.setFontSize(10);
			XWPFParagraph para1 = nTableheaderRow.getCell(1).getParagraphs().get(0);
			XWPFRun nTableRun1 = para1.createRun();
			nTableRun1.setBold(true);
			nTableRun1.setText("CVS Project Team:");
			nTableRun1.setFontFamily("Arial");
			nTableRun1.setFontSize(10);
			for (int j = 0; j < rows; j++) {
				XWPFTableRow nextRow = nTable.getRow(j + 1);
				nextRow.getCell(0).setText(
						j < vendorNameList.size() ? vendorNameList.get(j) + " - " + vendorRoleList.get(j) : "");
				nextRow.getCell(1)
						.setText(j < CVSNameList.size() ? CVSNameList.get(j) + " - " + CVSRoleList.get(j) : "");
			}

			int found = -1;
			for (int i = 0; i < tables.size(); i++) {
				XWPFTable table = tables.get(i);
				if (table.getText().contains("Vendor")) {
					found = i;
//					log.info(Double.toString(found));
					break;
				}
			}
			// Remove the old table
			if (found != -1) {
				doc.removeBodyElement(found - 1);
				doc.setTable(found, nTable);
			}

			// Create a new table
			List<String> months = Arrays.asList("January", "February", "March", "April", "May", "June", "July",
					"August", "September", "October", "November", "December");

			int tableIndex = -1;
			for (int i = 0; i < tables.size(); i++) {
				XWPFTable table = tables.get(i);
				if (table.getText().contains("Deliverable")) {
					tableIndex = i;
//					log.info(Double.toString(tableIndex));
					break;
				}
			}

			Date StartDate = user.getStartDate();
			SimpleDateFormat StartDateFormat = new SimpleDateFormat("MMMM");
			String startDate = StartDateFormat.format(StartDate);

			String endDate = "";
			if (user.getEndDate() != null) {
				Date EndDate = user.getEndDate();
				SimpleDateFormat EndDateFormat = new SimpleDateFormat("MMMM");
				endDate = EndDateFormat.format(EndDate);
			}
			List<String> RoleMonths = new ArrayList<String>();
			List<String> AllRoles = new ArrayList<String>();
			List<String> RoleLocations = new ArrayList<String>();
			List<Number> RoleRate = new ArrayList<Number>();
			HashMap<String, Integer> resourceCount = new HashMap<String, Integer>();
			List[][] RoleTotal = new List[100][100];
			int monthStartIndex = months.indexOf(startDate);
			int monthEndIndex = months.indexOf(endDate);

			for (int i = monthStartIndex; i <= monthEndIndex; i++) {
				RoleMonths.add(months.get(i));
			}

			for (Map<String, Object> RoleIterator : rowsData) {
				AllRoles.add((String) RoleIterator.get("CVS Role"));
				RoleLocations.add((String) RoleIterator.get("On/Off"));
				RoleRate.add((Number) RoleIterator.get("Grandfathered /CVS Rate"));
//	    	RoleTotal.add((Number) RoleIterator.get("Sat Jul 01 00:00:00 IST 2023"));
			}
			Date year = user.getStartDate();
			SimpleDateFormat opFormat = new SimpleDateFormat("yyyy");
			String yearString = opFormat.format(year);
			
			for (Map<String, Object> RoleIterator : rowsData) {
			    String role = (String) RoleIterator.get("CVS Role");
			    String location = (String) RoleIterator.get("On/Off");
			    Double rate = (Double) RoleIterator.get("Grandfathered /CVS Rate");
			    String key = role + "," + location + "," + rate.toString();
			    if (resourceCount.containsKey(key)) {
			        int count = resourceCount.get(key);
			        resourceCount.put(key, count + 1);
			    } else {
			        resourceCount.put(key, 1);
			    }
			}

			for (int i = 0; i < RoleMonths.size(); i++) {
				for (int j = 0; j < AllRoles.size(); j++) {
					RoleTotal[i][j] = new ArrayList();
					for (Map<String, Object> RoleIterator : rowsData) {
						String dateString = RoleMonths.get(i);
						SimpleDateFormat inputFormat = new SimpleDateFormat("MMMM yyyy");
						SimpleDateFormat outputFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy");
						Date date = inputFormat.parse(dateString + " " + yearString);
						String outputDateString = outputFormat.format(date);
						RoleTotal[i][j].add(RoleIterator.get(outputDateString));
					}					
				}
			}
			int dupRoleCount=0;
			
			Set<String> uniqueInputRows = new TreeSet<String>();
			
			for (int i = 0; i < RoleMonths.size(); i++) {
				for (int j = 0; j < AllRoles.size(); j++) {
					 String[] values = {AllRoles.get(j), RoleLocations.get(j), RoleRate.get(j).toString()};
			         String resourceKey = String.join(",", values);
			         
			         if (RoleTotal[i][j].get(j) != null 
			        		 && resourceCount.get(resourceKey) != null) {
			        	 String key = RoleMonths.get(i) + "," + AllRoles.get(j) + "," + 
						         RoleLocations.get(j) + "," + RoleRate.get(j).toString();
			             dupRoleCount+=1;   
				
			         if (!uniqueInputRows.contains(key) )  {
			        	 uniqueInputRows.add(key);
				            
			         	}
			         }
				}
			}
			
			
			ExcelService.budgetAmount = 0.0;

			// Deliverables Table Creation
			XWPFTable DeliverableTable = doc.createTable(RoleMonths.size() + 1, 2);
			XWPFTableRow deliverablesTableheaderRow = DeliverableTable.getRow(0);
			deliverablesTableheaderRow.getCell(0).setColor("CCE1FD");
			deliverablesTableheaderRow.getCell(1).setColor("CCE1FD");
			XWPFParagraph Para = deliverablesTableheaderRow.getCell(0).getParagraphs().get(0);
			XWPFRun deliverablesheaderTableRun = Para.createRun();
			deliverablesheaderTableRun.setBold(true);
			deliverablesheaderTableRun.setText("Deliverables");
			deliverablesheaderTableRun.setFontFamily("Arial");
			deliverablesheaderTableRun.setFontSize(10);
			XWPFParagraph Para1 = deliverablesTableheaderRow.getCell(1).getParagraphs().get(0);
			XWPFRun deliverablesTableheaderRun1 = Para1.createRun();
			deliverablesTableheaderRun1.setBold(true);
			deliverablesTableheaderRun1.setText("Date to complete");
			deliverablesTableheaderRun1.setFontFamily("Arial");
			deliverablesTableheaderRun1.setFontSize(10);
			if (tableIndex != -1) {
				doc.removeBodyElement(tableIndex - 1);

				Date SDate = user.getStartDate();
				SimpleDateFormat SDateFormat = new SimpleDateFormat("MMMM");
				String sDate = SDateFormat.format(SDate);
				
				SimpleDateFormat yearFormat = new SimpleDateFormat("yyyy");
				String yearDate = yearFormat.format(SDate);
				int YearDate = Integer.parseInt(yearDate);

				String eDate = "";
				if (user.getEndDate() != null) {
					Date EndDate = user.getEndDate();
					SimpleDateFormat EndDateFormat = new SimpleDateFormat("MMMM");
					eDate = EndDateFormat.format(EndDate);
				}

				int mStartIndex = 0;
				mStartIndex = months.indexOf(sDate);
				int mEndIndex = 0;
				if (eDate.length() > 0) {
					mEndIndex = months.indexOf(eDate);
					for (int i = mStartIndex + 1, j = 0; i <= mEndIndex + 1; i++, j++) {
						// XWPFTableRow row = newTable.createRow();
						// XWPFTableCell cell1 = row.createCell();
						XWPFTableRow row = DeliverableTable.getRow(j + 1);

						LocalDate date = LocalDate.of(YearDate, i, 1);
						XWPFTableCell cell1 = row.getCell(0);
						cell1.setText("Services for the Month of " + months.get(i - 1) + " "
								+ date.with(TemporalAdjusters.lastDayOfMonth()).getYear());
						XWPFTableCell cell2 = row.getCell(1);

						cell2.setText(date.with(TemporalAdjusters.lastDayOfMonth()).getDayOfMonth() + "-"
								+ months.get(i - 1).substring(0, 3) + "-"
								+ date.with(TemporalAdjusters.lastDayOfMonth()).getYear());

					}
				} else {
					if (mStartIndex == 0) {
						XWPFTableRow row = DeliverableTable.getRow(1);
						XWPFTableCell cell1 = row.getCell(0);
						LocalDate date = LocalDate.of(LocalDate.now().getYear(), mStartIndex + 1, 1);
						cell1.setText("Services for the Month of " + months.get(0) + " "
								+ date.with(TemporalAdjusters.lastDayOfMonth()).getYear());
						XWPFTableCell cell2 = row.getCell(1);
						cell2.setText(date.with(TemporalAdjusters.lastDayOfMonth()).getDayOfMonth() + "-"
								+ months.get(0).substring(0, 3) + "-"
								+ date.with(TemporalAdjusters.lastDayOfMonth()).getYear());
					} else {
						XWPFTableRow row = DeliverableTable.getRow(1);
						XWPFTableCell cell1 = row.getCell(0);
						LocalDate date = LocalDate.of(LocalDate.now().getYear(), mStartIndex, 1);
						cell1.setText("Services for the Month of " + months.get(mStartIndex) + " "
								+ date.with(TemporalAdjusters.lastDayOfMonth()).getYear());
						XWPFTableCell cell2 = row.getCell(1);
						cell2.setText(date.with(TemporalAdjusters.lastDayOfMonth()).getDayOfMonth() + "-"
								+ months.get(mStartIndex - 1).substring(0, 3) + "-"
								+ date.with(TemporalAdjusters.lastDayOfMonth()).getYear());
					}
				}

				// Insert the new table at the same position as the old table
				doc.setTable(tableIndex, DeliverableTable);
			}

			int width = 8000;
			for (int i = 0; i < RoleMonths.size(); i++) {
				for (int j = 0; j < AllRoles.size(); j++) {
					Object value = RoleTotal[i][j].get(j);
					if (value instanceof Number) {
						ExcelService.budgetAmount += ((Number) value).doubleValue();
					}
				}
			}

			double number =  ExcelService.budgetAmount;
			NumberFormat numberFormat = new DecimalFormat("#,##0.00");
			String formattedNumber = numberFormat.format(number);
			
			// Months and Roll Table Creation
			XWPFTable Trail = doc.createTable(uniqueInputRows.size() +2, 6);

			// Set the width of each column to be equal
			Trail.getCTTbl().addNewTblPr().addNewTblW().setW(BigInteger.valueOf(width));

			// Add the header row
			XWPFTableRow hRow = Trail.getRow(0);
			hRow.getCell(0).setColor("CCE1FD");
			hRow.getCell(1).setColor("CCE1FD");
			hRow.getCell(2).setColor("CCE1FD");
			hRow.getCell(3).setColor("CCE1FD");
			hRow.getCell(4).setColor("CCE1FD");
			hRow.getCell(5).setColor("CCE1FD");

			XWPFParagraph paraForMonths = hRow.getCell(0).getParagraphs().get(0);
			XWPFRun hRun = paraForMonths.createRun();
			hRun.setBold(true);
			hRun.setText("Months");
			hRun.setFontFamily("Arial");
			hRun.setFontSize(10);

			XWPFParagraph paraForRoles = hRow.getCell(1).getParagraphs().get(0);
			XWPFRun hRowRun2 = paraForRoles.createRun();
			hRowRun2.setText("Roles");
			hRowRun2.setBold(true);
			hRowRun2.setFontFamily("Arial");
			hRowRun2.setFontSize(10);

			XWPFParagraph paraForLocation = hRow.getCell(2).getParagraphs().get(0);
			XWPFRun hRowRun3 = paraForLocation.createRun();
			hRowRun3.setText("Location");
			hRowRun3.setBold(true);
			hRowRun3.setFontFamily("Arial");
			hRowRun3.setFontSize(10);

			XWPFParagraph paraForRate = hRow.getCell(3).getParagraphs().get(0);
			XWPFRun hRowRun4 = paraForRate.createRun();
			hRowRun4.setText("Rate");
			hRowRun4.setBold(true);
			hRowRun4.setFontFamily("Arial");
			hRowRun4.setFontSize(10);
			
			XWPFParagraph paraForCount = hRow.getCell(4).getParagraphs().get(0);
			XWPFRun hRowRun5 = paraForCount.createRun();
			hRowRun5.setText("Resource count");
			hRowRun5.setBold(true);
			hRowRun5.setFontFamily("Arial");
			hRowRun5.setFontSize(10);
			CTTblWidth width1 = hRow.getCell(4).getCTTc().addNewTcPr().addNewTcW();
			width1.setType(STTblWidth.DXA);
			width1.setW(BigInteger.valueOf(500));

			XWPFParagraph paraForTotal = hRow.getCell(5).getParagraphs().get(0);
			XWPFRun hRowRun6 = paraForTotal.createRun();
			hRowRun6.setText("Total");
			hRowRun6.setBold(true);
			hRowRun6.setFontFamily("Arial");
			hRowRun6.setFontSize(10);

			   int x = 1;
				String prevMonth = "";
				Set<String> uniqueRows = new TreeSet<String>();
				
				for (int i = 0; i < RoleMonths.size(); i++) {
				    for (int j = 0; j < AllRoles.size(); j++) {
				    	 String[] values = {AllRoles.get(j), RoleLocations.get(j), RoleRate.get(j).toString()};
				         String resourceKey = String.join(",", values);
				         if (RoleTotal[i][j].get(j) != null &&RoleTotal[i][j].get(j) !=" $- " && resourceCount.get(resourceKey) != null) {
				             String key = RoleMonths.get(i) + "," + AllRoles.get(j) + "," + RoleLocations.get(j) + "," + RoleRate.get(j).toString();
				             
				             if (!uniqueRows.contains(key) )  {
				            uniqueRows.add(key);
				            // create a new row and add it to the table
				            XWPFTableRow nRow = Trail.getRow(x);
				            x += 1;
				            
				            if (RoleMonths.get(i).equals(prevMonth)) {
				                // Skip setting the value for this cell and merge it with the previous cell
				                XWPFTableCell cell = nRow.getCell(0);
				                CTTcPr tcPr = cell.getCTTc().addNewTcPr();
				                tcPr.addNewVMerge().setVal(STMerge.CONTINUE);
				            } else {
			                    // Check the number of cells to merge based on the length of AllRoles
			                    int numCellsToMerge = values.length;
			                    if (numCellsToMerge > 1) {
			                        // Merge the cells
			                        CTVMerge vMerge = CTVMerge.Factory.newInstance();
			                        vMerge.setVal(STMerge.RESTART);
			                        XWPFTableCell cell = nRow.getCell(0);
			                        CTTcPr tcPr = cell.getCTTc().addNewTcPr();
			                        tcPr.setVMerge(vMerge);
			                        for (int k = 1; k < numCellsToMerge; k++) {
			                            XWPFTableRow row = Trail.getRow(x + k - 1);
			                            if (row != null && row.getCell(0) != null) {
			                                row.getCell(0).getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
			                            }
			                        }
			                    }
			                    nRow.getCell(0).setText(RoleMonths.get(i));
			                }

			                nRow.getCell(1).setText(AllRoles.get(j));
			                nRow.getCell(2).setText(RoleLocations.get(j));
			                nRow.getCell(3).setText("$ " + String.valueOf(RoleRate.get(j)));
			                nRow.getCell(4).setText(resourceCount.get(resourceKey) == null ? "0" : String.valueOf(resourceCount.get(resourceKey)));
			                
			                double num = ((Double)RoleTotal[i][j].get(j)) * (resourceCount.get(resourceKey));
			                String fNumber = numberFormat.format(num);
			                nRow.getCell(5).setText(RoleTotal[i][j].get(j) == null ? " $- "
			                        : "$ " +  fNumber);
			                
//			                nRow.getCell(5).setText(RoleTotal[i][j].get(j) == null ? " $- "
//			                        : "$ " + String.format("%.2f", ((Double)RoleTotal[i][j].get(j)) * (resourceCount.get(resourceKey))));
			                
			                prevMonth = RoleMonths.get(i);
			            }
			        }
			    }
			}
				XWPFTableRow nRow = Trail.getRow(uniqueInputRows.size() +1);
				nRow.getCell(0).setColor("CCE1FD");
				nRow.getCell(1).setColor("CCE1FD");
				nRow.getCell(2).setColor("CCE1FD");
				nRow.getCell(3).setColor("CCE1FD");
				nRow.getCell(4).setColor("CCE1FD");
				nRow.getCell(5).setColor("CCE1FD");
				
				XWPFParagraph paraForTotalInLastTableRow = hRow.getCell(4).getParagraphs().get(0);
				XWPFRun nRowRun = paraForTotalInLastTableRow.createRun();
				nRowRun.setText("Total");
				nRowRun.setBold(true);
				nRowRun.setFontFamily("Arial");
				hRowRun6.setFontSize(12);
				

				nRow.getCell(4).setText("Total");
				nRow.getCell(5).setText("$ " + formattedNumber);
	
//				nRow.getCell(5).setText("$ " + String.format("%.2f", ExcelService.budgetAmount));
				
				CTVerticalJc vAlign = CTVerticalJc.Factory.newInstance();
				vAlign.setVal(STVerticalJc.CENTER);
				for (int i = 1; i < Trail.getRows().size(); i++) {
				    Trail.getRow(i).getCell(0).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
				    Trail.getRow(i).getCell(0).getCTTc().addNewTcPr().setVAlign(vAlign);
				}  
			int targetFound = -1;
			for (int i = 0; i < tables.size(); i++) {
				XWPFTable Monthtable = tables.get(i);
				if (Monthtable.getText().contains("Month")) {
					targetFound = i;
//					log.info(Double.toString(targetFound));
					break;
				}
			}
			// Remove the old table
			if (targetFound != -1) {
				doc.removeBodyElement(targetFound - 1);
				doc.setTable(targetFound, Trail);
			}

			int tableCount = tables.size();
			XWPFTable lastTable = tables.get(tableCount - 1);
			XWPFTable last_2_Table = tables.get(tableCount - 2);
			XWPFTable last_3_Table = tables.get(tableCount - 3);
			doc.removeBodyElement(doc.getPosOfTable(lastTable));
			doc.removeBodyElement(doc.getPosOfTable(last_2_Table));
			doc.removeBodyElement(doc.getPosOfTable(last_3_Table));

			List<XWPFParagraph> paragraphs = doc.getParagraphs();
			for (XWPFParagraph paragraph : paragraphs) {
				List<XWPFRun> runs = paragraph.getRuns();
				for (XWPFRun run : runs) {
					String text = run.getText(0);
					for (Map<String, Object> row : rowsData) {
						ExcelService.sowName = (String) row.get("Contract#");
						if (text != null && text.contains("Contract")) {
							text = text.replace("Contract", (String) row.get("Contract#"));
							run.setText(text, 0);
						}

						if (text != null && text.contains("sow_effective_date")) {

//							String dateString = (String) row.get("Start Date");
							Date dateString = user.getStartDate();
//							log.info(dateString);
//							System.out.println(row.get("Start Date"));
							SimpleDateFormat inputFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy");
							SimpleDateFormat outputFormat = new SimpleDateFormat("MMMM dd, yyyy");
//							Date date = inputFormat.parse(dateString);
							String outputStartDateString = outputFormat.format(dateString);

							text = text.replace("sow_effective_date", outputStartDateString);
							run.setText(text, 0);
						}
						if (text != null && text.contains("year")) {

							String dateString = (String) row.get("End Date");
							SimpleDateFormat inputFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy");
							SimpleDateFormat outputFormat = new SimpleDateFormat("yyyy");
							Date date = inputFormat.parse(dateString);
							String outputYearString = outputFormat.format(date);

							text = text.replace("year", outputYearString);
							run.setText(text, 0);
						}

						if (text != null && text.contains("sow_end_date")) {

//							String dateString = (String) row.get("End Date");
							Date dateString = user.getEndDate();
							SimpleDateFormat inputFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy");
							SimpleDateFormat outputFormat = new SimpleDateFormat("MMMM dd, yyyy");
//							dateString = inputFormat.parse(dateString);
							String outputEndDateString = outputFormat.format(dateString);

							text = text.replace("sow_end_date", outputEndDateString);
							run.setText(text, 0);
						}
						if (text != null && text.contains("budget_amount")) {
							text = text.replace("budget_amount",
									"$" + formattedNumber);
							run.setText(text, 0);
						}
						break;
					}
				}
			}

			XWPFTable tbl = tables.get(3);
			for (XWPFTableRow row : tbl.getRows()) {
				for (XWPFTableCell cell : row.getTableCells()) {
					for (XWPFParagraph p : cell.getParagraphs()) {
						for (XWPFRun r : p.getRuns()) {
							String text = r.getText(0);
							if (text.contains("Budget_amount")) {
								text = text.replaceAll("Budget_amount",
										formattedNumber);
								r.setText(text, 0);
							}

						}
					}
				}
			}

			defaultName = user.getSowName();
//			String NewSowName = new String();	  

			ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
			doc.write(byteArrayOutputStream);
			HttpHeaders headers = new HttpHeaders();
			headers.setContentType(MediaType
					.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
			headers.add("Content-Disposition", "attachment; filename=SOW_Document1.docx");
			headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");
			ResponseEntity<byte[]> response = new ResponseEntity<byte[]>(byteArrayOutputStream.toByteArray(), headers,
					HttpStatus.OK);
			byteArrayOutputStream.close();
			return response;
		} else {
//			log.info("Application Failed to download the files, Kindly Please check the Date,Sheet Name and Cost Center values!!!!! ");
			return new ResponseEntity("Enter proper date", HttpStatus.BAD_REQUEST);

		}
	}

	public List FileNames() {
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

		for (int i = 0; i < ExcelService.defaultName.length(); i++) {
			NewSFName += ExcelService.defaultName.charAt(i);

			if (i == 2) {
				NewSFName += "SF";
			}
		}

		for (int i = 0; i < NewSFName.length(); i++) {
			NeWSFName += NewSFName.charAt(i);

			if (i == 5) {
				NeWSFName += ExcelService.sowName;
			}
		}
		NeWSFName += ".xlsx";

		List fileNames = new ArrayList<>();
		fileNames.add(NewSowName);
		fileNames.add(NeWSFName);
		NewSowName = "";
		NeWSFName = "";
		NewSFName = "";
		return fileNames;
	}

	public ResponseEntity<byte[]> insertDataIntoExcel(User user) throws Exception {

		FileInputStream excel_file = new FileInputStream(new File(submiTPath.getSubmitPath()));
		XSSFWorkbook workbook = new XSSFWorkbook(excel_file);

		// for sow submission file
		Sheet sheet = workbook.getSheet("SOW Submission Form");

		List sowNames = new ArrayList<>();
		sowNames = FileNames();

		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() == 956801) {
					cell.setCellValue(user.getEmpId());
				}
				if (cell.getCellType() == CellType.STRING && cell.getStringCellValue()
						.equals("COG2023-0XX.01_CCCC86_SOW_Business for Active Health-ChangeOrderForm#3.docx")) {
					cell.setCellValue(sowNames.get(0).toString());
				}

				if (cell.getCellType() == CellType.STRING && cell.getStringCellValue()
						.equals("2. Is this the first SOW we have signed for this client?")) {
					int cellIndex = cell.getColumnIndex();
					Cell nextCell = row.getCell(cellIndex + 1);
					nextCell.setCellValue("Yes");
				}

				if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {

					Date date = cell.getDateCellValue();

					SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
					String formattedDate = dateFormat.format(date);

					if (formattedDate.equals("01-03-2023")) {
						SimpleDateFormat logDateFormat = new SimpleDateFormat("dd-MM-yyyy");
						String newDate = logDateFormat.format(user.getStartDate());
						Date parsedDate = logDateFormat.parse(newDate);
						cell.setCellValue(parsedDate);
					}

					if (formattedDate.equals("31-12-2023")) {
						SimpleDateFormat logDateFormat = new SimpleDateFormat("dd-MM-yyyy");
						String newDate = logDateFormat.format(user.getEndDate());
						Date parsedDate = logDateFormat.parse(newDate);
						cell.setCellValue(parsedDate);
					}

				}

				if (cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() == 24026.4) {

					double cellValue = ExcelService.budgetAmount;
					cell.setCellValue(cellValue);

				}

				if (cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() == 383065) {

					Map<String, Map<String, String>> vendorMap = vendorData.getVendor();

					for (String vendorId : vendorMap.keySet()) {
						Map<String, String> vendorDetails = vendorMap.get(vendorId);
						String name = vendorDetails.get("name");
						String role = vendorDetails.get("role");

						if (role.equals("Client Relationship Manager")) {
							cell.setCellValue(vendorId);

						}
					}

				}

				if (cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() == 424844) {

					Map<String, Map<String, String>> vendorMap = vendorData.getVendor();

					for (String vendorId : vendorMap.keySet()) {
						Map<String, String> vendorDetails = vendorMap.get(vendorId);
						String name = vendorDetails.get("name");
						String role = vendorDetails.get("role");

						if (role.equals("Engagement Manager")) {
							cell.setCellValue(vendorId);

						}
					}

				}

			}
		}

		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		workbook.write(outputStream);
		HttpHeaders headers = new HttpHeaders();
		headers.setContentType(
				MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
		headers.add("Content-Disposition", "attachment; filename=Submit_File1.xlsx");
		headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");
		ResponseEntity<byte[]> response = new ResponseEntity<byte[]>(outputStream.toByteArray(), headers,
				HttpStatus.OK);
		outputStream.close();
		return response;
	}

	public List<Map<String, Object>> ReadAmendmentData(MultipartFile file, Amendment user) throws Exception {
		Workbook workbook = new XSSFWorkbook(file.getInputStream());
		Sheet sheet = workbook.getSheet(user.getSheetName());
		String columnName = "Cost Center";
		String AmendmentFilterColumn = "SOW Signed";
		int AmendmentFilterColumnIndex = -1;
		int columnIndex = -1;
		Row headerRow = sheet.getRow(10);
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		int count = 0;
		for (Cell cell : headerRow) {
			if (cell.getCellType() == CellType.STRING) {
				;
				if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
					columnIndex = cell.getColumnIndex();
				}
				if (cell.getStringCellValue().equalsIgnoreCase(AmendmentFilterColumn)) {
					AmendmentFilterColumnIndex = cell.getColumnIndex();
				}

			} else if (cell.getCellType() == CellType.NUMERIC) {
				if (Double.toString(cell.getNumericCellValue()).equalsIgnoreCase(columnName)) {
					columnIndex = cell.getColumnIndex();
				}
			}
		}

		for (int i = 0; i <= rowCount; i++) {
			Row row = sheet.getRow(i);
			if (row != null) {
				Cell cell = row.getCell(columnIndex);
				if (cell != null) {
					String cellValue = cell.getStringCellValue();
					if (cellValue.equals(user.getCostCenter())) {
						count++;
					}
				}
			}
		}

		ExcelService.amendmentCount = count;
		List<Row> rows = new ArrayList<Row>();
		for (int i = 11; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			Cell cell = row.getCell(columnIndex);
			Cell SOWSignedCell = row.getCell(AmendmentFilterColumnIndex);
			if (cell != null) {
				Cell cell1 = row.getCell(columnIndex + 1);
				Cell cell2 = row.getCell(columnIndex + 2);

				if (cell1.getDateCellValue() != null) {
					Date uDate = user.getStartDate();
					Date eDate = cell1.getDateCellValue();
					SimpleDateFormat outputFormat = new SimpleDateFormat("yyyy");
					String outputInputDateString = outputFormat.format(uDate);
					String outputExcelDateString = outputFormat.format(eDate);
					SimpleDateFormat oFormat = new SimpleDateFormat("dd-MM-yyyy");
					Date givenDate = user.getResourceDate(); // change the date as needed
					String oDString = oFormat.format(eDate);
					if (outputExcelDateString.equals(outputInputDateString)) {
						
//						Filter Based on Amendment Start Date
						
//						if (cell.getStringCellValue().equalsIgnoreCase(user.getCostCenter())
//								&& (oDString.equals(oFormat.format(givenDate))) || eDate.after(givenDate)) {
//							rows.add(row);
//						}
						
//						Filter Based on SOW Signed column
						
						if (cell.getStringCellValue().equalsIgnoreCase(user.getCostCenter()) 
								&& SOWSignedCell.getStringCellValue().equalsIgnoreCase("No")) {
							rows.add(row);
						}
					}

				}
			} else {
				break;
			}
		}

		List<Map<String, Object>> rowsData = new ArrayList<Map<String, Object>>();
		for (Row row : rows) {
			Map<String, Object> rowData = new HashMap<String, Object>();
			for (Cell cell : row) {
				int columnIndexx = cell.getColumnIndex();
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
					if (cell.getCellStyle().getDataFormat() == HSSFDataFormat
							.getBuiltinFormat("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)")) {
						rowData.put(columnNamee, "$" + cellValue.getNumberValue());
					} else {
						switch (evaluator.evaluateFormulaCell(cell)) {
						case NUMERIC:
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
		System.out.println(rowsData);
		return rowsData;
	}

	public ResponseEntity<byte[]> insertDataIntoAmendmentWord(MultipartFile file, Amendment user) throws Exception {
		FileInputStream inputStream = new FileInputStream(new File(amendmentPath.getAmendmentPath()));
		XWPFDocument doc = new XWPFDocument(inputStream);
		List<Map<String, Object>> rowsData = new ArrayList<Map<String, Object>>();
		rowsData = ReadAmendmentData(file, user);
		
		List<String> CVSTeamList = user.getCvsTeam();
		List<String> CVSNameList = new ArrayList<String>();

		Map<String, Map<String, String>> cvsMap = cvsData.getCvs();
		for (String cvsId : CVSTeamList) {
			Map<String, String> cvsDetails = cvsMap.get(cvsId);
			String cvsName = cvsDetails.get("name");
			CVSNameList.add(cvsName);

		}
		
		
		if (rowsData.size() > 0) {

			List<XWPFParagraph> paragraphs = doc.getParagraphs();
			for (XWPFParagraph paragraph : paragraphs) {
				List<XWPFRun> runs = paragraph.getRuns();
				for (XWPFRun run : runs) {
					String text = run.getText(0);
					for (Map<String, Object> row : rowsData) {
						ExcelService.sowName = (String) row.get("Contract#");
						if (text != null && text.contains("contract")) {
							text = text.replace("contract", (String) row.get("Contract#"));
							run.setText(text, 0);
						}

						if (text != null && text.contains("costCenter")) {
							text = text.replace("costCenter", "CC" + user.getCostCenter());
							run.setText(text, 0);
						}

						if (text != null && text.contains("amendmentStartDate")) {

//							String dateString = (String) row.get("Start Date");
							Date dateString = user.getStartDate();
//							log.info(dateString);
//							System.out.println(row.get("Start Date"));
							SimpleDateFormat inputFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy");
							SimpleDateFormat outputFormat = new SimpleDateFormat("MMMM dd, yyyy");
//							Date date = inputFormat.parse(dateString);
							String outputStartDateString = outputFormat.format(dateString);
							text = text.replace("amendmentStartDate", outputStartDateString);
							run.setText(text, 0);
						}
						
						if (text != null && text.contains("Powell, Ed")) {
							text = text.replace("Powell, Ed", CVSNameList.get(0));
							run.setText(text, 0);
						}

					}
				}
			}

			List<XWPFTable> tables = doc.getTables();
			XWPFTable tbl = tables.get(0);
			List<XWPFTableRow> rows = tbl.getRows();

			int tableCount = tables.size();
			for (XWPFTableRow row : tbl.getRows()) {
				for (XWPFTableCell cell : row.getTableCells()) {
					for (XWPFParagraph p : cell.getParagraphs()) {
						for (XWPFRun r : p.getRuns()) {
							String text = r.getText(0);

							if (text.contains("amendmentStartDate")) {
								Date startDate = user.getStartDate();
								SimpleDateFormat inputFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy");
								SimpleDateFormat outputFormat = new SimpleDateFormat("MMMM dd, yyyy");
								String outputStartDateString = outputFormat.format(startDate);
								text = text.replace("amendmentStartDate", outputStartDateString);
								r.setText(text, 0);
							}

							if (text.contains("amendmentEndDate")) {

								Date dateString = user.getEndDate();
								SimpleDateFormat inputFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy");
								SimpleDateFormat outputFormat = new SimpleDateFormat("MMMM dd, yyyy");
								String outputEndDateString = outputFormat.format(dateString);
								text = text.replace("amendmentEndDate", outputEndDateString);
								r.setText(text, 0);

							}

							
						}
					}
				}
			}

			List<XWPFTable> tables2 = doc.getTables();
			XWPFTable tbl2 = tables2.get(2);
			List<XWPFTableRow> rows2 = tbl2.getRows();

			int tableCount2 = tables2.size();
			for (XWPFTableRow row : tbl2.getRows()) {
				for (XWPFTableCell cell : row.getTableCells()) {
					for (XWPFParagraph p : cell.getParagraphs()) {
						for (XWPFRun r : p.getRuns()) {
							String text = r.getText(0);

							if (text.contains("resourcesData")) {
								text = text.replace("resourcesData", user.getAdditionResource());
								r.setText(text, 0);
							}

						}
					}
				}
			}
			

			List<String> months = Arrays.asList("January", "February", "March", "April", "May", "June", "July",
					"August", "September", "October", "November", "December");
			Date StartDate = user.getStartDate();
			SimpleDateFormat StartDateFormat = new SimpleDateFormat("MMMM");
			String startDate = StartDateFormat.format(StartDate);

			String endDate = "";
			if (user.getEndDate() != null) {
				Date EndDate = user.getEndDate();
				SimpleDateFormat EndDateFormat = new SimpleDateFormat("MMMM");
				endDate = EndDateFormat.format(EndDate);
			}
			List<String> RoleMonths = new ArrayList<String>();
			List<String> AllRoles = new ArrayList<String>();
			List<String> RoleLocations = new ArrayList<String>();
			List<String> RoleSkills = new ArrayList<String>();
			List<Number> RoleRate = new ArrayList<Number>();
			HashMap<String, Integer> resourceCount = new HashMap<String, Integer>();
			List[][] RoleTotal = new List[100][100];
			int monthStartIndex = months.indexOf(startDate);
			int monthEndIndex = months.indexOf(endDate);

			for (int i = monthStartIndex; i <= monthEndIndex; i++) {
				RoleMonths.add(months.get(i));
			}

			for (Map<String, Object> RoleIterator : rowsData) {
				AllRoles.add((String) RoleIterator.get("CVS Role"));
				RoleLocations.add((String) RoleIterator.get("On/Off"));
				RoleRate.add((Number) RoleIterator.get("Grandfathered /CVS Rate"));
				RoleSkills.add((String) RoleIterator.get("Skills"));
//    	RoleTotal.add((Number) RoleIterator.get("Sat Jul 01 00:00:00 IST 2023"));
			}
			Date year = user.getStartDate();
			SimpleDateFormat opFormat = new SimpleDateFormat("yyyy");
			String yearString = opFormat.format(year);
			
			for (Map<String, Object> RoleIterator : rowsData) {
			    String role = (String) RoleIterator.get("CVS Role");
			    String location = (String) RoleIterator.get("On/Off");
			    Double rate = (Double) RoleIterator.get("Grandfathered /CVS Rate");
			    String key = role + "," + location + "," + rate.toString();
			    if (resourceCount.containsKey(key)) {
			        int count = resourceCount.get(key);
			        resourceCount.put(key, count + 1);
			    } else {
			        resourceCount.put(key, 1);
			    }
			}


			for (int i = 0; i < RoleMonths.size(); i++) {
				for (int j = 0; j < AllRoles.size(); j++) {
					RoleTotal[i][j] = new ArrayList();
					for (Map<String, Object> RoleIterator : rowsData) {
						String dateString = RoleMonths.get(i);
						SimpleDateFormat inputFormat = new SimpleDateFormat("MMMM yyyy");
						SimpleDateFormat outputFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy");
						Date date = inputFormat.parse(dateString + " " + yearString);
						String outputDateString = outputFormat.format(date);
						RoleTotal[i][j].add(RoleIterator.get(outputDateString));
					}
				}
			}
			
			int dupRoleCount=0;
			
			Set<String> uniqueInputRows = new TreeSet<String>();
			
			for (int i = 0; i < RoleMonths.size(); i++) {
				for (int j = 0; j < AllRoles.size(); j++) {
					 String[] values = {AllRoles.get(j), RoleLocations.get(j), RoleRate.get(j).toString()};
			         String resourceKey = String.join(",", values);
			         
			         if (RoleTotal[i][j].get(j) != null 
			        		 && resourceCount.get(resourceKey) != null) {
			        	 String key = RoleMonths.get(i) + "," + AllRoles.get(j) + "," + 
						         RoleLocations.get(j) + "," + RoleRate.get(j).toString();
			             dupRoleCount+=1;   
				
			         if (!uniqueInputRows.contains(key) )  {
			        	 uniqueInputRows.add(key);
				            
			         	}
			         }
				}
			}
			
			
			ExcelService.budgetAmount = 0.0;
			int width = 9000;
			for (int i = 0; i < RoleMonths.size(); i++) {
				for (int j = 0; j < AllRoles.size(); j++) {
					Object value = RoleTotal[i][j].get(j);
					if (value instanceof Number) {
						ExcelService.budgetAmount += ((Number) value).doubleValue();
					}
				}
			}
			
			double number =  ExcelService.budgetAmount;
			NumberFormat numberFormat = new DecimalFormat("#,##0.00");
			String formattedNumber = numberFormat.format(number);

			// Months and Roll Table Creation
			XWPFTable Trail = doc.createTable(uniqueInputRows.size() +2, 7);

			// Set the width of each column to be equal
			Trail.getCTTbl().addNewTblPr().addNewTblW().setW(BigInteger.valueOf(width));

			// Add the header row
			XWPFTableRow hRow = Trail.getRow(0);
			hRow.getCell(0).setColor("CCE1FD");
			hRow.getCell(1).setColor("CCE1FD");
			hRow.getCell(2).setColor("CCE1FD");
			hRow.getCell(3).setColor("CCE1FD");
			hRow.getCell(4).setColor("CCE1FD");
			hRow.getCell(5).setColor("CCE1FD");
			hRow.getCell(6).setColor("CCE1FD");
			
			XWPFParagraph paraForMonths = hRow.getCell(0).getParagraphs().get(0);
			XWPFRun hRun = paraForMonths.createRun();
			hRun.setBold(true);
			hRun.setText("Months");
			hRun.setFontFamily("Arial");
			hRun.setFontSize(10);

			XWPFParagraph paraForRoles = hRow.getCell(1).getParagraphs().get(0);
			XWPFRun hRowRun2 = paraForRoles.createRun();
			hRowRun2.setText("Role-Level");
			hRowRun2.setBold(true);
			hRowRun2.setFontFamily("Arial");
			hRowRun2.setFontSize(10);

			XWPFParagraph paraForSkill = hRow.getCell(2).getParagraphs().get(0);
			XWPFRun hRowRun6 = paraForSkill.createRun();
			hRowRun6.setText("Skill");
			hRowRun6.setBold(true);
			hRowRun6.setFontFamily("Arial");
			hRowRun6.setFontSize(10);

			XWPFParagraph paraForLocation = hRow.getCell(3).getParagraphs().get(0);
			XWPFRun hRowRun3 = paraForLocation.createRun();
			hRowRun3.setText("Location");
			hRowRun3.setBold(true);
			hRowRun3.setFontFamily("Arial");
			hRowRun3.setFontSize(10);

			XWPFParagraph paraForRate = hRow.getCell(4).getParagraphs().get(0);
			XWPFRun hRowRun4 = paraForRate.createRun();
			hRowRun4.setText("Rate");
			hRowRun4.setBold(true);
			hRowRun4.setFontFamily("Arial");
			hRowRun4.setFontSize(10);

			XWPFParagraph paraForCount = hRow.getCell(5).getParagraphs().get(0);
			XWPFRun hRowRun5 = paraForCount.createRun();
			hRowRun5.setText("Resource count");
			hRowRun5.setBold(true);
			hRowRun5.setFontFamily("Arial");
			hRowRun5.setFontSize(10);
			CTTblWidth width1 = hRow.getCell(4).getCTTc().addNewTcPr().addNewTcW();
			width1.setType(STTblWidth.DXA);
			width1.setW(BigInteger.valueOf(500));

			XWPFParagraph paraForTotal = hRow.getCell(6).getParagraphs().get(0);
			XWPFRun hRowRun7 = paraForTotal.createRun();
			hRowRun7.setText("Total");
			hRowRun7.setBold(true);
			hRowRun7.setFontFamily("Arial");
			hRowRun7.setFontSize(10);
			
			 int x = 1;
			String prevMonth = "";
			Set<String> uniqueRows = new TreeSet<String>();
			
			for (int i = 0; i < RoleMonths.size(); i++) {
			    for (int j = 0; j < AllRoles.size(); j++) {
			    	 String[] values = {AllRoles.get(j), RoleLocations.get(j), RoleRate.get(j).toString()};
			         String resourceKey = String.join(",", values);
			         if (RoleTotal[i][j].get(j) != null &&RoleTotal[i][j].get(j) !=" $- " && resourceCount.get(resourceKey) != null) {
			             String key = RoleMonths.get(i) + "," + AllRoles.get(j) + "," + RoleLocations.get(j) + "," + RoleRate.get(j).toString();
			             
			             if (!uniqueRows.contains(key) )  {
			            uniqueRows.add(key);
			            // create a new row and add it to the table
			            XWPFTableRow nRow = Trail.getRow(x);
			            x += 1;
			            
			            if (RoleMonths.get(i).equals(prevMonth)) {
			                // Skip setting the value for this cell and merge it with the previous cell
			                XWPFTableCell cell = nRow.getCell(0);
			                CTTcPr tcPr = cell.getCTTc().addNewTcPr();
			                tcPr.addNewVMerge().setVal(STMerge.CONTINUE);
			            } else {
		                    // Check the number of cells to merge based on the length of AllRoles
		                    int numCellsToMerge = values.length;
		                    if (numCellsToMerge > 1) {
		                        // Merge the cells
		                        CTVMerge vMerge = CTVMerge.Factory.newInstance();
		                        vMerge.setVal(STMerge.RESTART);
		                        XWPFTableCell cell = nRow.getCell(0);
		                        CTTcPr tcPr = cell.getCTTc().addNewTcPr();
		                        tcPr.setVMerge(vMerge);
		                        for (int k = 1; k < numCellsToMerge; k++) {
		                            XWPFTableRow row = Trail.getRow(x + k - 1);
		                            if (row != null && row.getCell(0) != null) {
		                                row.getCell(0).getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
		                            }
		                        }
		                    }
		                    nRow.getCell(0).setText(RoleMonths.get(i));
		                }
					prevMonth = RoleMonths.get(i);
					nRow.getCell(1).setText(AllRoles.get(j));
					nRow.getCell(2).setText(RoleSkills.get(j));
					nRow.getCell(3).setText(RoleLocations.get(j));
					nRow.getCell(4).setText("$ " + String.valueOf(RoleRate.get(j)));
					 nRow.getCell(5).setText(resourceCount.get(resourceKey) == null ?
		                		"0" : String.valueOf(resourceCount.get(resourceKey)));
					 
					 double num = ((Double)RoleTotal[i][j].get(j)) * (resourceCount.get(resourceKey));
		                String fNumber = numberFormat.format(num);
		                nRow.getCell(6).setText(RoleTotal[i][j].get(j) == null ? " $- "
		                        : "$ " +  fNumber);
					 
//					 nRow.getCell(6).setText(RoleTotal[i][j].get(j) == null ? " $- "
//		                        : "$ " + String.format("%.2f", ((Double)RoleTotal[i][j].get(j)) * (resourceCount.get(resourceKey))));
		                
			             }
			         }
			    }
			}     
			
			XWPFTableRow nRow = Trail.getRow(uniqueInputRows.size() +1);
			nRow.getCell(0).setColor("CCE1FD");
			nRow.getCell(1).setColor("CCE1FD");
			nRow.getCell(2).setColor("CCE1FD");
			nRow.getCell(3).setColor("CCE1FD");
			nRow.getCell(4).setColor("CCE1FD");
			nRow.getCell(5).setColor("CCE1FD");
			nRow.getCell(6).setColor("CCE1FD");
			
			XWPFParagraph paraForTotalInLastTableRow = hRow.getCell(5).getParagraphs().get(0);
			XWPFRun nRowRun = paraForTotalInLastTableRow.createRun();
			nRowRun.setText("Total");
			nRowRun.setBold(true);
			nRowRun.setFontFamily("Arial");
			hRowRun6.setFontSize(12);
			

			nRow.getCell(5).setText("Total");
			nRow.getCell(6).setText("$ " + formattedNumber);
			
			CTVerticalJc vAlign = CTVerticalJc.Factory.newInstance();
			vAlign.setVal(STVerticalJc.CENTER);
			for (int i = 1; i < Trail.getRows().size(); i++) {
				Trail.getRow(i).getCell(0).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
				Trail.getRow(i).getCell(0).getCTTc().addNewTcPr().setVAlign(vAlign);
			}

			int targetFound = -1;
			for (int i = 0; i < tables.size(); i++) {
				XWPFTable Monthtable = tables.get(i);
				if (Monthtable.getText().contains("Month")) {
					targetFound = i;
//				log.info(Double.toString(targetFound));
					break;
				}
			}

			// Remove the old table
			if (targetFound != -1) {
				doc.removeBodyElement(targetFound - 1);
				doc.setTable(targetFound, Trail);
			}

			int tableCount1 = tables.size();
			XWPFTable lastTable = tables.get(tableCount1 - 1);
			doc.removeBodyElement(doc.getPosOfTable(lastTable));
			
			
			XWPFTable table = doc.getTables().get(1);

			for (XWPFTableRow row : table.getRows()) {
			    for (XWPFTableCell cell : row.getTableCells()) {
			        for (XWPFParagraph p : cell.getParagraphs()) {
			            for (XWPFRun r : p.getRuns()) {
			                String text = r.getText(0);
		                if (text != null && text.contains("count")) {
			                    text = text.replaceAll("count", String.format("%d", (int) user.getCcCount()));
			                    r.setText(text, 0);
			                }
			                
			                if (text != null && text.contains("amendmentStartdate")) {
								Date startDate1 = user.getStartDate();
								SimpleDateFormat inputFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy");
								SimpleDateFormat outputFormat = new SimpleDateFormat("MMMM dd, yyyy");
								String outputStartDateString = outputFormat.format(startDate1);
								text = text.replace("amendmentStartdate", outputStartDateString);
								r.setText(text, 0);
							}

							if (text != null && text.contains("amendmentEndDate")) {

								Date dateString = user.getEndDate();
							SimpleDateFormat inputFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy");
								SimpleDateFormat outputFormat = new SimpleDateFormat("MMMM dd, yyyy");
								String outputEndDateString = outputFormat.format(dateString);
								text = text.replace("amendmentEndDate", outputEndDateString);
								r.setText(text, 0);

							}
							
							
							if (text != null && text.contains("sowAmount")) {
							    double sowAmount = user.getSowAmount();
							    String fNumber = numberFormat.format(sowAmount);
							    text = text.replace("sowAmount", "$ "+fNumber);
							    r.setText(text, 0);
							}

							if (text != null && text.contains("amendmentAmount")) {
							    text = text.replace("amendmentAmount", "$ "+formattedNumber);
							    r.setText(text, 0);
							}
							double amount = user.getSowAmount() + ExcelService.budgetAmount;
							String fNumber = numberFormat.format(amount);
							if (text != null && text.contains("totalBudgetAmount")) {
							    text = text.replace("totalBudgetAmount", "$ "+fNumber);
							    r.setText(text, 0);
							}

			            }
			        }
			    }
			}

			defaultName = user.getAmendmentName();

			ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
			doc.write(byteArrayOutputStream);
			HttpHeaders headers = new HttpHeaders();
			headers.setContentType(MediaType
					.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
			headers.add("Content-Disposition", "attachment; filename=AMENDMENT_Document1.docx");
			headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");
			ResponseEntity<byte[]> response = new ResponseEntity<byte[]>(byteArrayOutputStream.toByteArray(), headers,
					HttpStatus.OK);
			byteArrayOutputStream.close();
			return response;
		}

		else {
			return new ResponseEntity("Enter proper date", HttpStatus.BAD_REQUEST);

		}

	}

	
	
	public ResponseEntity<byte[]> insertDataIntoAmendmentExcel(Amendment user) throws Exception {
		FileInputStream excel_file = new FileInputStream(new File(submiTPath.getSubmitPath()));
		XSSFWorkbook workbook = new XSSFWorkbook(excel_file);

		// for sow submission file
		Sheet sheet = workbook.getSheet("SOW Submission Form");

		List sowNames = new ArrayList<>();
		sowNames = FileNames();

		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() == 956801) {
					cell.setCellValue(user.getEmpId());
				}
				if (cell.getCellType() == CellType.STRING && cell.getStringCellValue()
						.equals("COG2023-0XX.01_CCCC86_SOW_Business for Active Health-ChangeOrderForm#3.docx")) {
					cell.setCellValue(sowNames.get(0).toString());
				}

				if (cell.getCellType() == CellType.STRING && cell.getStringCellValue()
						.equals("2. Is this the first SOW we have signed for this client?")) {
					int cellIndex = cell.getColumnIndex();
					Cell nextCell = row.getCell(cellIndex + 1);
					nextCell.setCellValue("No");
				}

				if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().equals("2. Type of document:")) {
					int cellIndex = cell.getColumnIndex();
					Cell nextCell = row.getCell(cellIndex + 1);
					nextCell.setCellValue("SOW Amendment");
				}

				if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {

					Date date = cell.getDateCellValue();

					SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
					String formattedDate = dateFormat.format(date);

					if (formattedDate.equals("01-03-2023")) {
						SimpleDateFormat logDateFormat = new SimpleDateFormat("dd-MM-yyyy");
						String newDate = logDateFormat.format(user.getStartDate());
						Date parsedDate = logDateFormat.parse(newDate);
						cell.setCellValue(parsedDate);
					}

					if (formattedDate.equals("31-12-2023")) {
						SimpleDateFormat logDateFormat = new SimpleDateFormat("dd-MM-yyyy");
						String newDate = logDateFormat.format(user.getEndDate());
						Date parsedDate = logDateFormat.parse(newDate);
						cell.setCellValue(parsedDate);
					}

				}

				if (cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() == 24026.4) {

					double cellValue = ExcelService.totalBudgetAmount;
					cell.setCellValue(cellValue);

				}

				if (cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() == 383065) {

					Map<String, Map<String, String>> vendorMap = vendorData.getVendor();

					for (String vendorId : vendorMap.keySet()) {
						Map<String, String> vendorDetails = vendorMap.get(vendorId);
						String name = vendorDetails.get("name");
						String role = vendorDetails.get("role");

						if (role.equals("Client Relationship Manager")) {
							cell.setCellValue(vendorId);

						}
					}

				}

				if (cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() == 424844) {

					Map<String, Map<String, String>> vendorMap = vendorData.getVendor();

					for (String vendorId : vendorMap.keySet()) {
						Map<String, String> vendorDetails = vendorMap.get(vendorId);
						String name = vendorDetails.get("name");
						String role = vendorDetails.get("role");

						if (role.equals("Engagement Manager")) {
							cell.setCellValue(vendorId);

						}
					}

				}

			}
		}

		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		workbook.write(outputStream);
		HttpHeaders headers = new HttpHeaders();
		headers.setContentType(
				MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
		headers.add("Content-Disposition", "attachment; filename=Submit_File1.xlsx");
		headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");
		ResponseEntity<byte[]> response = new ResponseEntity<byte[]>(outputStream.toByteArray(), headers,
				HttpStatus.OK);
		outputStream.close();
		return response;
	}
}