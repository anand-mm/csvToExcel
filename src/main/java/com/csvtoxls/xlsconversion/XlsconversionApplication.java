package com.csvtoxls.xlsconversion;

import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@SpringBootApplication
@RestController
public class XlsconversionApplication {

	public static void main(String[] args) {
		SpringApplication.run(XlsconversionApplication.class, args);
	}

	@Autowired
	private static HttpServletResponse response;

	@GetMapping("/csv")
	public static void csvToXLSX() {
		try {
			String csvFileAddress = "C:\\Users\\USER\\Downloads\\test_csv1.csv";
			String fileName = FilenameUtils.getBaseName(csvFileAddress); // csv file address
			String xlsxFileAddress = fileName + ".xlsx"; // xlsx file address
			HSSFWorkbook workBook = new HSSFWorkbook();
			HSSFSheet sheet = workBook.createSheet("sheet1");
			String currentLine = null;
			int RowNum = 0;
			BufferedReader br = new BufferedReader(new FileReader(csvFileAddress));
			while ((currentLine = br.readLine()) != null) {
				String str[] = currentLine.split(",");
				RowNum++;
				HSSFRow currentRow = sheet.createRow(RowNum);
				for (int i = 0; i < str.length; i++) {
					currentRow.createCell(i).setCellValue(str[i]);
				}
			}
			FileOutputStream fileOutputStream = new FileOutputStream(xlsxFileAddress);
			workBook.write(fileOutputStream);
			fileOutputStream.close();
			br.close();
			workBook.close();
			System.out.println("Done");

		} catch (Exception ex) {
			System.out.println(ex.getMessage() + "Exception in try");
		}
	}

}
