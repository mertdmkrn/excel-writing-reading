import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import excel.Read;
import excel.Write;
import excel.model.Personal;

public class Main {

	public static void main(String[]args) throws IOException {
		
		Personal personal1 = new Personal(1,"Mert Demirkıran","Software",5000);
		Personal personal2 = new Personal(2,"Emrah Gökden","Marketing",3500);
		Personal personal3 = new Personal(3,"Kadir Memiş","Accounting",5000);
		Personal personal4 = new Personal(4,"Enes Kaya","Security",3500);
		
		List<Personal> personalList = Arrays.asList(personal1,personal2,personal3,personal4);
		
		Write excelWrite = new Write();
		
		excelWrite.writeExcel(personalList, "personalList.xlsx");
		
		System.out.println("Excel file write...");
		
		System.out.println("Excel file read...");
		
		Read excelRead = new Read();
		
		System.out.println(excelRead.readingExcel("personalList.xlsx"));
		
	}
}
