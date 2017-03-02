package com.excel.demo;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;



import com.excel.demo.pojo.CoachingReportsJSON;
import com.excel.demo.pojo.Reports;

public class WriteExcelDemo {

	/** Header Section Variables */
	private static final Integer STARTING_ROW_FOR_REPORT_HEADER_NAME = 0;

	private static final Integer STARTING_ROW_FOR_COACHING_SESSIONS_JSON = 18;
	private static final Integer STARTING_COLUMN_FOR_COACHING_SESSIONS_JSON = 1;
	
	private static final Integer TOTAL_REPORTS_IN_ONE_ROW = 6;

	public static void main(String[] args) {
		// Blank workbook
		HSSFWorkbook workbook = new HSSFWorkbook();

		// Create a blank sheet
		HSSFSheet sheet = workbook.createSheet("Employee Data");
		/** #################### Write Header ################################## */
		writeReportHeaderName("Top Action Report", sheet);

		/**
		 * #################### Write CoachingReportsJSON
		 * ##################################
		 */
		CoachingReportsJSON dataCoachingSessionsBasicJson = createCoachingReportsJSON();
		writeCoachingReportsJSON(dataCoachingSessionsBasicJson, sheet);

		/**
		 * #################### Write CoachingSessionsJson
		 * ##################################
		 */
		Map<String, Object[]> dataCoachingSessionsJson = createCoachingSessionsJson();
		addCoachingSessionsJson(dataCoachingSessionsJson, sheet);

		try {
			// Write the workbook in file system
			writeToFile("C:/Users/796412/Desktop/Excel/howtodoinjava_demo.xls",
					workbook);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static void writeCoachingReportsJSON(
			CoachingReportsJSON dataCoachingSessionsBasicJson, HSSFSheet sheet) {
		/**
		 * ############### Coaching Sessions Header ######################
		 * **/		
		setCellValueAtLocation(sheet, 2, 2, "Coaching Sessions", getBoldFont(sheet.getWorkbook()));
		
		/**
		 * ############### Coaching Sessions Header values ######################
		 * **/
		setCellValueAtLocation(sheet, 3, 2, dataCoachingSessionsBasicJson.getTotalForCoach()+"", null);
		setCellValueAtLocation(sheet, 3, 3, dataCoachingSessionsBasicJson.getCoachName()+"", null);
		
		setCellValueAtLocation(sheet, 4, 2, dataCoachingSessionsBasicJson.getTotalForOthers()+"", null);
		setCellValueAtLocation(sheet, 4, 3, dataCoachingSessionsBasicJson.getSelectedRoleName()+"", null);
		
		setCellValueAtLocation(sheet, 5, 2, dataCoachingSessionsBasicJson.getTotal()+"", null);
		setCellValueAtLocation(sheet, 5, 3, dataCoachingSessionsBasicJson.getTotalCaption()+"", getBoldFont(sheet.getWorkbook()));	
		
		/**
		 * ############### Coaching Actions Header ######################
		 * **/		
		setCellValueAtLocation(sheet, 2, 5, "Coaching Actions", getBoldFont(sheet.getWorkbook()));
		
		int totalReportsListCount=dataCoachingSessionsBasicJson.getReportsList().size();
		int totalReportsListColumnCount=(totalReportsListCount/TOTAL_REPORTS_IN_ONE_ROW)+((totalReportsListCount%TOTAL_REPORTS_IN_ONE_ROW)>0?1:0);
		
		List<Reports> listReports=dataCoachingSessionsBasicJson.getReportsList();
		for(int i=0;i<totalReportsListColumnCount;i++){
			
			for(int j=0;j<TOTAL_REPORTS_IN_ONE_ROW;j++){
				int index=i*TOTAL_REPORTS_IN_ONE_ROW+j;
				if(index<totalReportsListCount){
					Reports r=listReports.get(index);
//					setCellValueAtLocation(sheet, 3+j, 5, r.getPercentage()+"%", null);
//					setCellValueAtLocation(sheet, 3+j, 6, r.getActionName()+"", null);
					setCellValueAtLocation(sheet, 3+j, 5+i, r.getPercentage()+"%", null);
					setCellValueAtLocation(sheet, 3+j, 6+i, r.getActionName()+"", null);
				}
				
			}
			
		}

	}

	private static CoachingReportsJSON createCoachingReportsJSON() {
		CoachingReportsJSON objCoachingReportsJSON = new CoachingReportsJSON();

		objCoachingReportsJSON.setCoachName("Jeff Achtermann");
		objCoachingReportsJSON.setTotalForCoach(89);
		objCoachingReportsJSON.setSelectedRoleName("Quality Managers");
		objCoachingReportsJSON.setTotalForOthers(432);
		objCoachingReportsJSON.setTotalCaption("Total");
		objCoachingReportsJSON.setTotal(999);

		List<Reports> reportsList = new ArrayList<Reports>();
		reportsList.add(new Reports("VNA Training required", 10));
		reportsList.add(new Reports("VNA Training required", 10));
		reportsList.add(new Reports("VNA Training required", 10));
		reportsList.add(new Reports("VNA Training required", 10));
		reportsList.add(new Reports("VNA Training required", 10));

		reportsList.add(new Reports("VNA Training required", 10));
		reportsList.add(new Reports("VNA Training required", 10));
		reportsList.add(new Reports("VNA Training required", 10));
		reportsList.add(new Reports("VNA Training required", 10));
		reportsList.add(new Reports("VNA Training required", 10));
		objCoachingReportsJSON.setReportsList(reportsList);

		return objCoachingReportsJSON;
	}

	private static void writeReportHeaderName(String string, HSSFSheet sheet) {

		Row row = sheet.createRow((short) 0);
		Cell cell = row.createCell((short) 0);
		cell.setCellValue(string);

		sheet.addMergedRegion(new CellRangeAddress(
				0, // first row (0-based)
				1, // last row (0-based)
				0, // first column (0-based)
				1 // last column (0-based)
		));
	}

	private static void addCoachingSessionsJson(Map<String, Object[]> data,
			HSSFSheet sheet) {
		// Iterate over data and write to sheet
		Set<String> keyset = data.keySet();
		int rownum = STARTING_ROW_FOR_COACHING_SESSIONS_JSON;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object[] objArr = data.get(key);
			int cellnum = STARTING_COLUMN_FOR_COACHING_SESSIONS_JSON;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof String)
					cell.setCellValue((String) obj);
				else if (obj instanceof Integer)
					cell.setCellValue((Integer) obj);
			}
		}
	}

	private static Map<String, Object[]> createCoachingSessionsJson() {
		// This data needs to be written (Object[])
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[] { "Date", "Agent", "Coach", "Feedback" });
		data.put("2", new Object[] { "06/25/2015 02:00 PM", "Amit", "Shukla",
				"Lorem ipsum dolor sit amet, consectetur adipiscing" });
		data.put("3", new Object[] { "06/25/2015 02:00 PM", "Lokesh", "Gupta",
				"Lorem ipsum dolor sit amet, consectetur adipiscing" });
		data.put("4", new Object[] { "06/25/2015 02:00 PM", "John", "Adwards",
				"Lorem ipsum dolor sit amet, consectetur adipiscing" });
		data.put("5", new Object[] { "06/25/2015 02:00 PM", "Brian", "Schultz",
				"Lorem ipsum dolor sit amet, consectetur adipiscing" });

		return data;
	}

	private static void writeToFile(String url, HSSFWorkbook workbook)
			throws IOException {
		FileOutputStream out = new FileOutputStream(new File(url));
		workbook.write(out);
		out.close();
		System.out
				.println("howtodoinjava_demo.xlsx written successfully on disk.");
	}
	
	private static CellStyle getBoldFont(HSSFWorkbook workbook){
		CellStyle style = workbook.createCellStyle();
	    Font font = workbook.createFont();
	    font.setBoldweight(Font.BOLDWEIGHT_BOLD);
	    style.setFont(font);
	    return style;
	}
	
	private static Cell setCellValueAtLocation(HSSFSheet sheet,int rowIndex,int colIndex,String value,CellStyle style){
		Row row = sheet.createRow((short) rowIndex);
		Cell cell = row.createCell((short) colIndex);
		cell.setCellValue(value);
		if(style!=null){
			cell.setCellStyle(style);
		}
		
		return cell;
	}
}