package com.wptexcelintegration;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

@Component
public class ReadExcelHelper {

	public List<Url> readExcelFile(String filePath) {
		try {
			FileInputStream excelFile = new FileInputStream(new File(filePath));
			Workbook workbook = new XSSFWorkbook(excelFile);

			Sheet sheet = workbook.getSheet("RequestSheet");
			Iterator<Row> rows = sheet.iterator();

			List<Url> lstUrl = new ArrayList<Url>();

			int rowNumber = 0;
			while (rows.hasNext()) {
				Row currentRow = (Row) rows.next();

				if (rowNumber == 0) {
					rowNumber++;
					continue;
				}

				Iterator<Cell> cellsInRow = currentRow.iterator();

				Url url = new Url();

				int cellIndex = 0;
				while (cellsInRow.hasNext()) {
					Cell currentCell = (Cell) cellsInRow.next();

					if (cellIndex == 0) { // URL
						url.setUrl(String.valueOf(currentCell.getStringCellValue()));
					} else if (cellIndex == 1) { // KEY
						url.setKey(currentCell.getStringCellValue());

					} else if (cellIndex == 2) { // Location
						url.setLocation(currentCell.getStringCellValue());
					}

					cellIndex++;
				}

				lstUrl.add(url);
			}

			workbook.close();

			return lstUrl;
		} catch (IOException e) {
			throw new RuntimeException("FAIL = " + e.getMessage());
		}
	}

}
