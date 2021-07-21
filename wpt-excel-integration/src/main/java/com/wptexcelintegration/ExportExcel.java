package com.wptexcelintegration;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import com.fasterxml.jackson.databind.JsonNode;

@Service
public class ExportExcel {

	int rowCount = 0;

	public void writeObjectsToExcelFile(JsonNode fetchedData) throws IOException {

		String filePath = "D:\\Book1.xlsx";

		File file = new File(filePath);

		try {
			XSSFWorkbook workbook = null;

			if (!file.exists()) {
				workbook = new XSSFWorkbook();
			}

			else {
				try (FileInputStream fI = new FileInputStream(file)) {
					workbook = new XSSFWorkbook(fI);

				}
			}

			XSSFSheet spreadsheet = workbook.getSheet("RequestSheet");
			if (workbook.getNumberOfSheets() != 0) {
				for (int i = 1; i < workbook.getNumberOfSheets(); i++) {
					if (workbook.getSheetName(i).equals("ResultSheet")) {
						spreadsheet = workbook.getSheet("ResultSheet");
					} else
						spreadsheet = workbook.createSheet("ResultSheet");
				}
			} else {
				spreadsheet = workbook.createSheet("ResultSheet");
			}

			String[] COLUMNs = { "Test-URL", "Test-Id", "Total Blocking Time (TBT)", "Largest Contentful Paint(LCP)",
					"Cumulative Layout Shift (CLS)", "SpeedIndex", "Time To FirstByte" };

			Font headerFont = workbook.createFont();
			headerFont.setBold(true);
			headerFont.setColor(IndexedColors.BLUE.getIndex());

			CellStyle headerCellStyle = workbook.createCellStyle();
			headerCellStyle.setFont(headerFont);

			Row headerRow = spreadsheet.createRow(0);

			for (int col = 0; col < COLUMNs.length; col++) {
				Cell cell = headerRow.createCell(col);
				cell.setCellValue(COLUMNs[col]);
				cell.setCellStyle(headerCellStyle);
			}

			XSSFRow row = (XSSFRow) spreadsheet.createRow(++rowCount);

			XSSFCell cell1 = row.createCell(0);
			cell1.setCellValue(fetchedData.get("URL").toString());

			XSSFCell cell2 = row.createCell(1);
			cell2.setCellValue(fetchedData.get("testID").toString());

			XSSFCell cell3 = row.createCell(2);
			cell3.setCellValue(fetchedData.get("TotalBlockingTime").toString());

			XSSFCell cell4 = row.createCell(3);
			cell4.setCellValue(fetchedData.get("chromeUserTiming.LargestContentfulPaint").toString());

			XSSFCell cell5 = row.createCell(4);
			cell5.setCellValue(fetchedData.get("chromeUserTiming.CumulativeLayoutShift").toString());

			XSSFCell cell6 = row.createCell(5);
			cell6.setCellValue(fetchedData.get("SpeedIndex").toString());

			XSSFCell cell7 = row.createCell(6);
			cell7.setCellValue(fetchedData.get("TTFB").toString());

			FileOutputStream outputStream = new FileOutputStream(file);
			workbook.write(outputStream);

		} catch (UnsupportedOperationException e) {
			e.printStackTrace();
		}

		System.out.println("Exported excel successfully");

	}

}
