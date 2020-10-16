package com.example.demo;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;

@SpringBootApplication
public class DemoApachePoiApplication {

	private static final Logger log = LoggerFactory.getLogger(DemoApachePoiApplication.class);
	private String _ARROW = "---> ";
	
	public static void main(String[] args) {
//		SpringApplication.run(DemoApachePoiApplication.class, args);
		
		System.exit(
			SpringApplication.exit(
				SpringApplication.run(DemoApachePoiApplication.class, args)));
	}

	
	
	@Bean
	public CommandLineRunner demo() {
		return (args) -> {
			
			log.info(String.format("%s Start -------------------------", _ARROW));
			
			
			Workbook workbook = new XSSFWorkbook();
			
			Sheet sheet = workbook.createSheet("TestSheet");
			sheet.setColumnWidth(0, 6000);
			sheet.setColumnWidth(1, 6000);
			
			Row header = sheet.createRow(0);
			
			CellStyle headerStyle = workbook.createCellStyle();
			headerStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
			headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			 
			XSSFFont font = ((XSSFWorkbook) workbook).createFont();
			font.setFontName("Arial");
			font.setFontHeightInPoints((short) 12);
			font.setBold(true);
			headerStyle.setFont(font);
			
			Cell headerCell = header.createCell(0);
			headerCell.setCellValue("Name");
			headerCell.setCellStyle(headerStyle);
			 
			headerCell = header.createCell(1);
			headerCell.setCellValue("Age");
			headerCell.setCellStyle(headerStyle);
			
			CellStyle style = workbook.createCellStyle();
			style.setWrapText(true);
			 
			Row row = sheet.createRow(1);
			Cell cell = row.createCell(0);
			cell.setCellValue("John Smith");
			cell.setCellStyle(style);
			 
			cell = row.createCell(1);
			cell.setCellValue(20);
			cell.setCellStyle(style);
			
			File currDir = new File("C:/Users/ulixe/Documents/Luca/02_SPINDOX/MyDemo");
			String path = currDir.getAbsolutePath();
			String fileLocation = path + "/temp.xlsx";
			 
			log.info(String.format("%s file location: %s", _ARROW, fileLocation));
			
			FileOutputStream outputStream = new FileOutputStream(fileLocation);
			workbook.write(outputStream);
			workbook.close();
			
			
			log.info(String.format("%s Stop  -------------------------", _ARROW));
		};
	}
	
}
