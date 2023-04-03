package pack;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class Excel{
	XSSFSheet sh;
	  
	public String readData(int i,int j) {
		  Row r=sh.getRow(i);
		  Cell c=r.getCell(j);
		  int celltype=c.getCellType();//0 numeric,1 string
		  switch(celltype) {
		  case Cell.CELL_TYPE_NUMERIC:
		  {
			  double a=c.getNumericCellValue();
			  return String.valueOf(a);
		  }
		  case Cell.CELL_TYPE_STRING:
		  {
			  return c.getStringCellValue();
		  }
		  }
		  return null;
	  }
	  
	   public Excel() throws IOException {
		FileInputStream f=new FileInputStream("C:\\Users\\HP\\Documents\\New.xlsx");
		XSSFWorkbook workbook=new XSSFWorkbook(f);
		XSSFSheet sh=workbook.getSheet("Sheet1");
		
	}
	
}



	

	


