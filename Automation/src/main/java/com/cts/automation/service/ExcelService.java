package com.cts.automation.service;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.math.BigInteger;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.temporal.TemporalAdjusters;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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
import org.apache.poi.xwpf.usermodel.BreakType;
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

	public List<Map<String, Object>> ReadBasedOnCondition(MultipartFile file, User user) throws Exception {
//		FileInputStream inputStream = (FileInputStream) file.getInputStream();;
//		Workbook workbook = new XSSFWorkbook(inputStream);
		Workbook workbook = new XSSFWorkbook(file.getInputStream());
//		Sheet sheet = workbook.getSheetAt(1);
		Sheet sheet = workbook.getSheet(user.getSheetName());
		String columnName = "Cost Center";
		int columnIndex = -1;
		Row headerRow = sheet.getRow(10);
		for (Cell cell : headerRow) {
			if (cell.getCellType() == CellType.STRING) {
				;
				if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
					columnIndex = cell.getColumnIndex();
				}

			} else if (cell.getCellType() == CellType.NUMERIC) {
				if (Double.toString(cell.getNumericCellValue()).equalsIgnoreCase(columnName)) {
					columnIndex = cell.getColumnIndex();
				}
			}
		}

		List<Row> rows = new ArrayList<Row>();
//		for(Row row : sheet) {
		for (int i = 11; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			Cell cell = row.getCell(columnIndex);
			if (cell != null) {
				Cell cell1 = row.getCell(columnIndex + 1);
				Cell cell2 = row.getCell(columnIndex + 2);

				if (cell1.getDateCellValue() != null) {
					Date uDate = user.getStartDate();
					Date eDate = cell1.getDateCellValue();
					Date endDate = user.getEndDate();
					Date ExcelEndDate = cell2.getDateCellValue();
					SimpleDateFormat outputFormat = new SimpleDateFormat("yyyy");
					String outputInputDateString = outputFormat.format(uDate);
					String outputExcelDateString = outputFormat.format(eDate);
					String outputUserEndDateString = outputFormat.format(endDate);
					String outputExcelEndDateString = outputFormat.format(ExcelEndDate);

					if (outputExcelDateString.equals(outputInputDateString)
							&& outputUserEndDateString.equals(outputExcelEndDateString)) {
						if (cell.getStringCellValue().equalsIgnoreCase(user.getCostCenter())) {
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
					if (cell.getCellStyle().getDataFormat() == HSSFDataFormat
							.getBuiltinFormat("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)")) {
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

	public ResponseEntity<byte[]> insertDataIntoWord(MultipartFile file, User user) throws Exception {
		FileInputStream inputStream = new FileInputStream(new File(
				"C:\\Users\\2066253\\repository\\Excel-Automation\\COG2023-0XX_CCCC86_SOW_Business for Active Health_2023 (4).docx"));
		XWPFDocument doc = new XWPFDocument(inputStream);
		List<Map<String, Object>> rowsData = new ArrayList<Map<String, Object>>();
		rowsData = ReadBasedOnCondition(file, user);
		if (rowsData.size() > 0) {

			List<XWPFTable> tables = doc.getTables();

			List<String> vendorTeamList = new ArrayList<String>();
			vendorTeamList = user.getVendorTeam();

			List<String> CVSTeamList = new ArrayList<String>();
			CVSTeamList = user.getCvsTeam();

			int rows = Math.max(vendorTeamList.size(), CVSTeamList.size());

			XWPFTable nTable = doc.createTable(rows + 1, 2);
			XWPFTableRow nTableheaderRow = nTable.getRow(0);
			XWPFParagraph para = nTableheaderRow.getCell(0).getParagraphs().get(0);
			XWPFRun nTableRun = para.createRun();
			nTableRun.setBold(true);
			nTableRun.setText("Vendor Project Team:");
			XWPFParagraph para1 = nTableheaderRow.getCell(1).getParagraphs().get(0);
			XWPFRun nTableRun1 = para1.createRun();
			nTableRun1.setBold(true);
			nTableRun1.setText("CVS Project Team:");
			for (int j = 0; j < rows; j++) {
				XWPFTableRow nextRow = nTable.getRow(j + 1);
				nextRow.getCell(0).setText(j < vendorTeamList.size() ? vendorTeamList.get(j) : "");
				nextRow.getCell(1).setText(j < CVSTeamList.size() ? CVSTeamList.get(j) : "");
			}

			int found = -1;
			for (int i = 0; i < tables.size(); i++) {
				XWPFTable table = tables.get(i);
				if (table.getText().contains("Vendor")) {
					found = i;
					log.info(Double.toString(found));
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
//			XWPFTable newTable = doc.createTable();
//			XWPFTableRow row1 = newTable.createRow();
//			XWPFTableCell cell0 = row1.createCell();
//			XWPFTableCell cell01 = row1.createCell();
//			XWPFParagraph para0 = cell0.getParagraphs().get(0);
//			XWPFRun run0 = para0.createRun();
//			run0.setBold(true);
//			run0.setText("Deliverable");
//
//			XWPFParagraph para01 = cell01.getParagraphs().get(0);
//			XWPFRun run01 = para01.createRun();
//			run01.setBold(true);
//			run01.setText("Date to be Complete");
//
			int tableIndex = -1;
			for (int i = 0; i < tables.size(); i++) {
				XWPFTable table = tables.get(i);
				if (table.getText().contains("Deliverable")) {
					tableIndex = i;
					log.info(Double.toString(tableIndex));
					break;
				}
			}
//			// Remove the old table
//			if (tableIndex != -1) {
//				doc.removeBodyElement(tableIndex - 1);
//
//				Date StartDate = user.getStartDate();
//				SimpleDateFormat StartDateFormat = new SimpleDateFormat("MMMM");
//				String startDate = StartDateFormat.format(StartDate);
//
//				String endDate = "";
//				if (user.getEndDate() != null) {
//					Date EndDate = user.getEndDate();
//					SimpleDateFormat EndDateFormat = new SimpleDateFormat("MMMM");
//					endDate = EndDateFormat.format(EndDate);
//				}
//
//				int monthStartIndex = 0;
//				monthStartIndex = months.indexOf(startDate);
//				int monthEndIndex = 0;
//				if (endDate.length() > 0) {
//					monthEndIndex = months.indexOf(endDate);
//					for (int i = monthStartIndex + 1; i <= monthEndIndex + 1; i++) {
//						XWPFTableRow row = newTable.createRow();
//						XWPFTableCell cell1 = row.createCell();
//						LocalDate date = LocalDate.of(LocalDate.now().getYear(), i, 1);
//						cell1.setText("Services for the Month of " + months.get(i - 1) + " "
//								+ date.with(TemporalAdjusters.lastDayOfMonth()).getYear());
//						XWPFTableCell cell2 = row.createCell();
//						cell2.setText(date.with(TemporalAdjusters.lastDayOfMonth()).getDayOfMonth() + "-"
//								+ months.get(i - 1).substring(0, 3) + "-"
//								+ date.with(TemporalAdjusters.lastDayOfMonth()).getYear());
//					}
//				} else {
//					if (monthStartIndex == 0) {
//						XWPFTableRow row = newTable.createRow();
//						XWPFTableCell cell1 = row.createCell();
//						LocalDate date = LocalDate.of(LocalDate.now().getYear(), monthStartIndex + 1, 1);
//						cell1.setText("Services for the Month of " + months.get(0) + " "
//								+ date.with(TemporalAdjusters.lastDayOfMonth()).getYear());
//						XWPFTableCell cell2 = row.createCell();
//						cell2.setText(date.with(TemporalAdjusters.lastDayOfMonth()).getDayOfMonth() + "-"
//								+ months.get(0).substring(0, 3) + "-"
//								+ date.with(TemporalAdjusters.lastDayOfMonth()).getYear());
//					} else {
//						XWPFTableRow row = newTable.createRow();
//						XWPFTableCell cell1 = row.createCell();
//						LocalDate date = LocalDate.of(LocalDate.now().getYear(), monthStartIndex, 1);
//						cell1.setText("Services for the Month of " + months.get(monthStartIndex) + " "
//								+ date.with(TemporalAdjusters.lastDayOfMonth()).getYear());
//						XWPFTableCell cell2 = row.createCell();
//						cell2.setText(date.with(TemporalAdjusters.lastDayOfMonth()).getDayOfMonth() + "-"
//								+ months.get(monthStartIndex - 1).substring(0, 3) + "-"
//								+ date.with(TemporalAdjusters.lastDayOfMonth()).getYear());
//					}
//				}
//
//				// Insert the new table at the same position as the old table
//				doc.setTable(tableIndex, newTable);
//			}

			// Months and Roll Table Creation
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
			List[][] RoleTotal = new List[100][100];
			int monthStartIndex = months.indexOf(startDate);
			int monthEndIndex = months.indexOf(endDate);

			for (int i = monthStartIndex; i <= monthEndIndex; i++) {
				RoleMonths.add(months.get(i));
			}

			for (Map<String, Object> RoleIterator : rowsData) {
				AllRoles.add((String) RoleIterator.get("CVS Role"));
				RoleLocations.add((String) RoleIterator.get("Location"));
				RoleRate.add((Number) RoleIterator.get("Grandfathered /CVS Rate"));
//	    	RoleTotal.add((Number) RoleIterator.get("Sat Jul 01 00:00:00 IST 2023"));
			}
			Date year = user.getEndDate();
			SimpleDateFormat opFormat = new SimpleDateFormat("yyyy");
			String yearString = opFormat.format(year);

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
			double budgetAmount = 0.0;


			// Deliverables Table Creation
			XWPFTable DeliverableTable = doc.createTable(RoleMonths.size() + 1, 2);
			XWPFTableRow deliverablesTableheaderRow = DeliverableTable.getRow(0);
			XWPFParagraph Para = deliverablesTableheaderRow.getCell(0).getParagraphs().get(0);
			XWPFRun deliverablesheaderTableRun = Para.createRun();
			deliverablesheaderTableRun.setBold(true);
			deliverablesheaderTableRun.setText("Deliverables");
			XWPFParagraph Para1 = deliverablesTableheaderRow.getCell(1).getParagraphs().get(0);
			XWPFRun deliverablesTableheaderRun1 = Para1.createRun();
			deliverablesTableheaderRun1.setBold(true);
			deliverablesTableheaderRun1.setText("Date to complete");
			if (tableIndex != -1) {
				doc.removeBodyElement(tableIndex - 1);

				Date SDate = user.getStartDate();
				SimpleDateFormat SDateFormat = new SimpleDateFormat("MMMM");
				String sDate = StartDateFormat.format(SDate);

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
					for (int i = mStartIndex + 1; i <= mEndIndex + 1; i++) {
						// XWPFTableRow row = newTable.createRow();
						// XWPFTableCell cell1 = row.createCell();
						XWPFTableRow row = DeliverableTable.getRow(i);

						LocalDate date = LocalDate.of(LocalDate.now().getYear(), i, 1);
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
			
			
			
			XWPFTable table = doc.createTable(RoleMonths.size() + 1, 5);

			// Set the width of each column to be equal
			int width = 8000;
			table.setWidth("100%");
			table.getCTTbl().addNewTblPr().addNewTblW().setW(BigInteger.valueOf(width));

			// Add the header row
			XWPFTableRow headerRow = table.getRow(0);
			headerRow.getCell(0).setText("Months");
			headerRow.getCell(1).setText("Roles");
			headerRow.getCell(2).setText("Location");
			headerRow.getCell(3).setText("Rate");
			headerRow.getCell(4).setText("Total");

			for (int i = 0; i < RoleMonths.size(); i++) {
				table.getRow(i + 1).getCell(0).setText(RoleMonths.get(i));

				for (int j = 0; j < AllRoles.size(); j++) {
					XWPFParagraph p = table.getRow(i + 1).getCell(1).addParagraph();
					p.createRun().setText(AllRoles.get(j));
					XWPFRun run = p.createRun();
					run.addBreak(BreakType.TEXT_WRAPPING);

					XWPFParagraph p1 = table.getRow(i + 1).getCell(2).addParagraph();
					p1.createRun().setText(RoleLocations.get(j));
					XWPFRun run1 = p1.createRun();
					run1.addBreak(BreakType.TEXT_WRAPPING);

					XWPFParagraph p2 = table.getRow(i + 1).getCell(3).addParagraph();
					p2.createRun().setText("$ " + String.valueOf(RoleRate.get(j)));
					XWPFRun run2 = p2.createRun();
					run2.addBreak(BreakType.TEXT_WRAPPING);

					XWPFParagraph p3 = table.getRow(i + 1).getCell(4).addParagraph();
					if (RoleTotal[i][j].get(j) == null) {
						p3.createRun().setText("--");
						XWPFRun run3 = p3.createRun();
						run3.addBreak(BreakType.TEXT_WRAPPING);
					} else {
						p3.createRun().setText("$ " + String.format("%.2f", RoleTotal[i][j].get(j)));
						Object value = RoleTotal[i][j].get(j);
						if (value instanceof Number) {
							budgetAmount += ((Number) value).doubleValue();
						}
						XWPFRun run3 = p3.createRun();
						run3.addBreak(BreakType.TEXT_WRAPPING);
					}
				}
			}

			int targetFound = -1;
			for (int i = 0; i < tables.size(); i++) {
				XWPFTable Monthtable = tables.get(i);
				if (Monthtable.getText().contains("Month")) {
					targetFound = i;
					log.info(Double.toString(targetFound));
					break;
				}
			}
			// Remove the old table
			if (targetFound != -1) {
				doc.removeBodyElement(targetFound - 1);
				doc.setTable(targetFound, table);
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
						if (text != null && text.contains("Contract")) {
							text = text.replace("Contract", (String) row.get("Contract#"));
							run.setText(text, 0);
						}

						if (text != null && text.contains("sow_effective_date")) {

							String dateString = (String) row.get("Start Date");
							log.info(dateString);
							System.out.println(row.get("Start Date"));
							SimpleDateFormat inputFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy");
							SimpleDateFormat outputFormat = new SimpleDateFormat("MMMM dd, yyyy");
							Date date = inputFormat.parse(dateString);
							String outputStartDateString = outputFormat.format(date);

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

							String dateString = (String) row.get("End Date");
							SimpleDateFormat inputFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy");
							SimpleDateFormat outputFormat = new SimpleDateFormat("MMMM dd, yyyy");
							Date date = inputFormat.parse(dateString);
							String outputEndDateString = outputFormat.format(date);

							text = text.replace("sow_end_date", outputEndDateString);
							run.setText(text, 0);
						}
						if (text != null && text.contains("budget_amount")) {
							text = text.replace("budget_amount", "$" + String.format("%.2f", budgetAmount));
							run.setText(text, 0);
						}
						break;
					}
				}
			}

//	    XWPFTable tbl = tables.get(3);
//	    for (XWPFTableRow row : tbl.getRows()) {
//	    for (XWPFTableCell cell : row.getTableCells()) {
//            for (XWPFParagraph p : cell.getParagraphs()) {
//                for (XWPFRun r : p.getRuns()) {
//                    String text = r.getText(0);
//                    if (text.contains("Budget_amount")) {
//                        text = text.replaceAll("Budget_amount",  "$"+String.format("%.2f", budgetAmount));
//                        r.setText(text, 0);
//                    }
//                }
//            }
//        }
//	    }

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
			return new ResponseEntity("Enter proper date", HttpStatus.BAD_REQUEST);

		}
	}
}
