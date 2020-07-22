package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read {

	public StringBuilder readingExcel(String excelFilePath){
		
		StringBuilder output = new StringBuilder(); 
		
		 try {

	            FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
	            Workbook workbook = new XSSFWorkbook(excelFile);
	            Sheet sheet = workbook.getSheetAt(0);
	            int rowLenght = sheet.getLastRowNum();
	            
	            for(int i=1;i<rowLenght;i++) {
	            	Row row = sheet.getRow(i);
	            	int cellLenght = row.getLastCellNum();
	            	for(int j=0;j<cellLenght;j++) {
	            		Cell currentCell = row.getCell(j);
	            		if (currentCell.getCellType() == CellType.STRING)
	            			output.append(currentCell.getStringCellValue()+"\t");
	            		else if (currentCell.getCellType() == CellType.NUMERIC)
	            			output.append((int)currentCell.getNumericCellValue()+"\t");
	            	}
	            	output.append("\n");
	            }
	            
	        } catch (FileNotFoundException e) {
	            e.printStackTrace();
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
		 
		 return output;

	    }
		
	}
