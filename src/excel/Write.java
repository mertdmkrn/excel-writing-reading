package excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import excel.model.Personal;

public class Write {
	
	
	public void writeExcel(List<Personal> personalList, String excelFilePath) throws IOException {
		
	    Workbook workbook = new XSSFWorkbook();
	    Sheet sheet = workbook.createSheet();
	    
	    createHeaderRow(sheet);
	    
	    int rowCount = 1;
	    sheet.setHorizontallyCenter(true);

	    for (Personal per: personalList) {
	        Row row = sheet.createRow(rowCount);
	        writeBook(per, row, sheet);
	        rowCount++;
	    }
	 
	    try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
	        workbook.write(outputStream);
	    }
	}
	
	public void writeBook(Personal personal, Row row,Sheet sheet) {
		
	    CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
	    Font font = sheet.getWorkbook().createFont();
	    font.setBold(false);
	    font.setFontHeightInPoints((short) 11);
	    cellStyle.setAlignment(HorizontalAlignment.CENTER);
	    font.setFontName("Arial");
	    cellStyle.setFont(font);	
		
	    Cell cell = row.createCell(0);
	    cell.setCellValue(personal.getPerId());
	    cell.setCellStyle(cellStyle);
	 
	    cell = row.createCell(1);
	    cell.setCellValue(personal.getPerName());
	    cell.setCellStyle(cellStyle);
	 
	    cell = row.createCell(2);
	    cell.setCellValue(personal.getPerDept());
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(3);
	    cell.setCellValue(personal.getPerSalary());
	    cell.setCellStyle(cellStyle);
	    
	}
	
	public void createHeaderRow(Sheet sheet) {
		 
	    CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
	    Font font = sheet.getWorkbook().createFont();
	    font.setBold(true);
	    font.setColor(IndexedColors.RED.getIndex());
	    font.setFontHeightInPoints((short) 11);
	    cellStyle.setAlignment(HorizontalAlignment.CENTER);
	    font.setFontName("Arial"); 
	    cellStyle.setFont(font);
	 
	    Row row = sheet.createRow(0);
	    
	    Cell cellId = row.createCell(0);
	    cellId.setCellStyle(cellStyle);
	    cellId.setCellValue("ID");
	 
	    Cell cellName = row.createCell(1);
	    cellName.setCellStyle(cellStyle);
	    cellName.setCellValue("NAME");
	 
	    Cell cellDept = row.createCell(2);
	    cellDept.setCellStyle(cellStyle);
	    cellDept.setCellValue("DEPARTMENT");
	    
	    Cell cellSalary = row.createCell(3);
	    cellSalary.setCellStyle(cellStyle);
	    cellSalary.setCellValue("SALARY");
	}
	
}
