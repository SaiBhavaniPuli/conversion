
package com.convert.ExcelToCsv;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import com.opencsv.CSVWriter;

@SpringBootApplication
public class ExcelToCsvApplication {

	public static void main(String[] args) throws IOException {

		
		 try {
		  
		  FileInputStream excelFile = new FileInputStream(new File("D://saiBhavani/test.xlsx"));
		  XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
		  XSSFSheet datatypeSheet = workbook.getSheetAt(0);
		  Iterator<Row> iterator = datatypeSheet.iterator();
		  
		  FileWriter filewriter=new FileWriter("D://saiBhavani/test1.csv");
		  CSVWriter csvwriter=new CSVWriter(filewriter); 
		 
		  while (iterator.hasNext()) {
		 
		  Row currentRow = iterator.next(); 
		 
		  Iterator<Cell> cellIterator = currentRow.iterator();
		  int i=0; 
		  String[] s=new String[20];
		 while (cellIterator.hasNext()) {
			 
		  Cell currentCell = cellIterator.next();
		 
		  switch(currentCell.getCellType())  
			{  
			case NUMERIC:   //field that represents numeric cell type  
			//getting the value of the cell as a number  
			System.out.print(currentCell.getNumericCellValue()+ "\t\t");  
			s[i++]=currentCell.getNumericCellValue()+"";
			break;  
			case STRING:    //field that represents string cell type  
			//getting the value of the cell as a string  
			System.out.print(currentCell.getStringCellValue()+ "\t\t");  
			s[i++]=currentCell.getStringCellValue();
			break;  
			case FORMULA:    //field that represents string cell type  
				//getting the value of the cell as a string  
				System.out.print(currentCell.getDateCellValue()+ "\t\t");  
				s[i++]=currentCell.getDateCellValue()+"";
				break;  
			}  
		 
		  } 
		 System.out.println();
		 csvwriter.writeNext(s);
		  
		  }csvwriter.close();
		  excelFile.close();}catch (FileNotFoundException e) {
		  e.printStackTrace(); }
		 catch (IOException e) { e.printStackTrace(); }
		 

		
	}

}
