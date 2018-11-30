package com.hughes.summary;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	
	@SuppressWarnings("unchecked")
	public static void main(String[] args) throws IOException, InvalidFormatException {

		String fileName = args[0];

		// Creating a Workbook from an Excel file (.xls or .xlsx)
	   Workbook workbook = WorkbookFactory.create(new File(fileName));
		
	   Workbook outputWorkBook = new XSSFWorkbook();
		  
       Map<String, Report> reportMap = new HashMap<>();

		// Retrieving the number of sheets in the Workbook
		System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

		/*
		 * ============================================================= Iterating over
		 * all the sheets in the workbook (Multiple ways)
		 * =============================================================
		 */
		
		//FileOutputStream fileOut = new FileOutputStream("/Users/anirvanroy/Documents/dev-workspace/summaryreport/src/main/resources/Output-report.xlsx");
		//String file = "/Users/anirvanroy/Documents/dev-workspace/summaryreport/src/main/resources/Output-report.xlsx";
		String file = "/Users/anirvanroy/Documents/workspace/runs/Output-report.xlsx";
		DataFormatter dataFormatter = new DataFormatter();
		for (Sheet sheet : workbook) {
			
			Set set = new TreeSet();
			Set totalSiteSet = new TreeSet();
			Map<String, Integer> map = new HashMap();
			Map<String, Set> mapSet = new HashMap();
		    if (sheet.getSheetName().equalsIgnoreCase("EntireBPData")
					|| sheet.getSheetName().equalsIgnoreCase("EnitreSonicData")) {
				System.out.println("********************************* "+sheet.getSheetName()+" *********************************");
				int rowCount = 0;
				for (Row row : sheet) {
					if (rowCount > 0) {
						int i = 0;
						String ownerId = "";
						for (Cell cell : row) {
							String cellValue = dataFormatter.formatCellValue(cell);
							if (i == 0) {
								ownerId = cellValue;
								set.add(cellValue);
								if ((sheet.getSheetName().equalsIgnoreCase("EntireBPData")
										&& !cellValue.equalsIgnoreCase("BPO"))
										|| (sheet.getSheetName().equalsIgnoreCase("EntireSonicData")
												&& !cellValue.equalsIgnoreCase("SON"))) {
									//System.out.println(ownerId);
									if (mapSet.get(ownerId) == null) {
										Set treeSet = new TreeSet();
										mapSet.put(ownerId, treeSet);
									}
								}

							}

							if (sheet.getSheetName().equalsIgnoreCase("EntireBPData") && i == 16) {
								totalSiteSet.add(cellValue);
								if (!ownerId.equalsIgnoreCase("BPO")) {
									Set<String> tempSet = mapSet.get(ownerId);
									tempSet.add(cellValue);
									mapSet.put(ownerId, tempSet);
								}

							}

							if (sheet.getSheetName().equalsIgnoreCase("EntireSonicData") && i == 15) {
								totalSiteSet.add(cellValue);

								if (!ownerId.equalsIgnoreCase("SON")) {
									Set tempSet = mapSet.get(ownerId);
									tempSet.add(cellValue);
									mapSet.put(ownerId, tempSet);
								}

							}

							i++;
						}
					}
					rowCount++;
				}

				Integer siteCount = 0;
				Integer ownerSiteCount = 0;
				Integer otherBigOwner = 0;
				Integer otherBigOwnerCount = 0;
				/*
				 * for(String key : map.keySet()) { if(map.get(key) >= 25) {
				 * System.out.println(key + " Count "+map.get(key)); siteCount += map.get(key);
				 * ownerSiteCount++; } }
				 */

				for (String key : mapSet.keySet()) {
					//System.out.println(key + " Count " + mapSet.get(key).size());
					if (mapSet.get(key).size() >= 25) {
						System.out.println(key + " Count " + mapSet.get(key).size());
						ownerSiteCount++;
						siteCount += mapSet.get(key).size();
					}
					
					if (mapSet.get(key).size() < 25) {
						System.out.println(key + " Count " + mapSet.get(key).size());
						otherBigOwner++;
						otherBigOwnerCount += mapSet.get(key).size();
					}

				}
				System.out.println("<------------------------- Summary Start ---------------------------->");
				System.out.println("Total No of sites " + totalSiteSet.size());
				System.out.println("No of Unique site owners " + set.size());
				System.out.println("No of Unique Big owners having >= 25 count " + ownerSiteCount);
				System.out.println("No of total sites with Big Owners having >= 25 count " + siteCount);
				System.out.println("No of Unique Others owners having < 25 count " + otherBigOwner);
				System.out.println("No of total sites with Big Owners having < 25 count " + otherBigOwnerCount);
				System.out.println("<------------------------- Summary End ---------------------------->");
				
				Report report = new Report();
				report.setTotalSite(totalSiteSet.size());
				report.setUniqueSiteOwners(set.size());
				report.setUniqueBigOwner(ownerSiteCount);
				report.setBigOwnerSiteCounts(siteCount);
				report.setOtherBigOwner(otherBigOwner);
				report.setOtherBigOwnerSiteCounts(otherBigOwnerCount);
				
				reportMap.put(sheet.getSheetName() + "_Summary_Report", report);
				
		        
			}
		}

		printCellValue(outputWorkBook, reportMap, file);
	    // Closing the workbook
		workbook.close();
	}

	private static void printCellValue(Workbook workbook, Map<String, Report> reportMap, String fileName) throws IOException {
		// Create a Sheet
		for(String key : reportMap.keySet()) {
			Report report = reportMap.get(key);
			 Sheet sheet = workbook.createSheet(key);
			 Row headerRow = sheet.createRow(0);
		        Cell headerCell = headerRow.createCell(0);
		        headerCell.setCellValue("Franchise List");
		        
		        String[] columns = {"Unique", "Sites"};

		        int rowNum = 1;
		        // Creating cells
		        int j=0;
		        Row row = sheet.createRow(rowNum);
		        for(int i = 1; i <= columns.length; i++) {
		            Cell cell = row.createCell(i);
		            cell.setCellValue(columns[j]);
		            j++;
		            
		        }
		        rowNum ++;
		        Row thirdRow = sheet.createRow(rowNum);
		        Cell thirdRowCell = thirdRow.createCell(0);
		        thirdRowCell.setCellValue("Total Landscape");
		        
		        Cell thirdRowCell2 = thirdRow.createCell(1);
		        thirdRowCell2.setCellValue(report.getUniqueSiteOwners());
		        
		        Cell thirdRowCell3 = thirdRow.createCell(2);
		        thirdRowCell3.setCellValue(report.getTotalSite());
		        
		        rowNum ++;
		        Row fourthRow = sheet.createRow(rowNum);
		        Cell fourthRowCell = fourthRow.createCell(0);
		        fourthRowCell.setCellValue("Corporate");
		        
		        Cell fourthRowCell2 = fourthRow.createCell(1);
		        fourthRowCell2.setCellValue("0");
		        
		        Cell fourthRowCell3 = fourthRow.createCell(2);
		        fourthRowCell3.setCellValue("0");
		        
		        rowNum++;
		        Row fifthRow = sheet.createRow(rowNum);
		        Cell fifthRowCell = fifthRow.createCell(0);
		        fifthRowCell.setCellValue("Influencers & Detractors");
		        
		        Cell fifthRowCell2 = fifthRow.createCell(1);
		        fifthRowCell2.setCellValue("");
		        
		        Cell fifthRowCell3 = fifthRow.createCell(2);
		        fifthRowCell3.setCellValue("");
		        
		        rowNum++;
		        Row sixthRow = sheet.createRow(rowNum);
		        Cell sixthRowCell = sixthRow.createCell(0);
		        sixthRowCell.setCellValue("Big Owners (25+ Sites)");
		        
		        Cell sixthRowCell2 = sixthRow.createCell(1);
		        sixthRowCell2.setCellValue(report.getUniqueBigOwner());
		        
		        Cell sixthRowCell3 = sixthRow.createCell(2);
		        sixthRowCell3.setCellValue(report.getBigOwnerSiteCounts());
		        
		        rowNum++;
		        Row seventhRow = sheet.createRow(rowNum);
		        Cell seventhRowCell = seventhRow.createCell(0);
		        seventhRowCell.setCellValue("Others");
		        
		        Cell seventhRowCell2 = seventhRow.createCell(1);
		        seventhRowCell2.setCellValue(report.getOtherBigOwner());
		        
		        Cell seventhRowCell3 = seventhRow.createCell(2);
		        seventhRowCell3.setCellValue(report.getOtherBigOwnerSiteCounts());
		}
		    FileOutputStream fileOut = new FileOutputStream(fileName);
		 	workbook.write(fileOut);
			fileOut.close();
			workbook.close();
       }
}