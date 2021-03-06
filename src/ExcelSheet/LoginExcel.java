
package ExcelSheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LoginExcel {


		//To Get the data from the excel sheet 
		public String getExcelData(String sheetname,int rownum,int cellnum) throws InvalidFormatException
		{
			String retval=null;
			try
			{
				FileInputStream fis = new FileInputStream("C:\\Orders\\orders.xlsx"); //give the path of excel sheet from where the data to be fetched
				Workbook wb=WorkbookFactory.create(fis);
				Sheet s= wb.getSheet(sheetname);
				Row r=s.getRow(rownum);
				Cell c = r.getCell(cellnum);
				
				retval = ((org.apache.poi.ss.usermodel.Cell) c).getStringCellValue();
			}
			catch(FileNotFoundException e)
			{
				e.printStackTrace();
			}
			catch(EncryptedDocumentException e)
			{
				e.printStackTrace();
			}
			catch(IOException e)
			{
				e.printStackTrace();
			}
			return retval;
		}
		
		
		//The method returns a zero based integer of the last row in the sheet  <which contains last data in that sheet>
		public int getRowCount(String sheetname) throws InvalidFormatException
		{
			
			int rowcount=0;
			try
			{
				FileInputStream fis= new FileInputStream("C:\\Orders\\orders.xlsx");
				Workbook wb=WorkbookFactory.create(fis);
				Sheet s= wb.getSheet(sheetname);
				rowcount = s.getLastRowNum();
				
			}
			catch(FileNotFoundException e)
			{
				e.printStackTrace();
			}
			catch(EncryptedDocumentException e)
			{
				e.printStackTrace();
			}
			catch(IOException e)
			{
				e.printStackTrace();
			}
			
			return rowcount;
		}
		//create page
		public void createPage(String sheetName) throws InvalidFormatException
		{
			try
			{
				FileInputStream fis = new FileInputStream("C:\\Orders\\orders.xlsx");
				XSSFWorkbook wb = new XSSFWorkbook(fis); 
				Sheet s= wb.createSheet("Organized");
				FileOutputStream fos=new FileOutputStream("C:\\Orders\\orders.xlsx");
				wb.write(fos);
				fos.close();
			}
			catch(FileNotFoundException e)
			{
				e.printStackTrace();
			}
			catch(EncryptedDocumentException e)
			{
				e.printStackTrace();
			}
			catch(IOException e)
			{
				e.printStackTrace();
			}
			
}
		//This method returns the data into the excel sheet
		public void setExcelData(String sheetname,int rownum,int cellnum,String data) throws InvalidFormatException
		{
			try
			{
				FileInputStream  fis = new FileInputStream ("C:\\Orders\\orders.xlsx");
				Workbook wb=WorkbookFactory.create(fis);
				Sheet s= wb.getSheet(sheetname);
				Row r=s.createRow(rownum);
				Cell c= r.createCell(cellnum);
				c.setCellValue((String)data);
				FileOutputStream fos=new FileOutputStream("C:\\Orders\\orders.xlsx");
//				fos.flush();
				wb.write(fos);
				fos.close();
			}
			
				catch(FileNotFoundException e)
				{
					e.printStackTrace();
				}
				catch(EncryptedDocumentException e)
				{
					e.printStackTrace();
				}
				catch(IOException e)
				{
					e.printStackTrace();
				}
				
	}
		
		public void setExcelData1(String sheetname,int rownum,int cellnum,ArrayList<String> data) throws InvalidFormatException
		{
			try
			{
				FileInputStream  fis = new FileInputStream ("C:\\Orders\\orders.xlsx");
				Workbook wb=WorkbookFactory.create(fis);
				Sheet s= wb.getSheet(sheetname);
				Row r=s.createRow(rownum);
				for(int i=0;i<data.size();i++)
				{
					Cell c= r.createCell(cellnum++);
					System.out.println(data.get(i));
					c.setCellValue((String)data.get(i));
					cellnum=cellnum++;
				}
				FileOutputStream fos=new FileOutputStream("C:\\Orders\\orders.xlsx");
//				fos.flush();
				wb.write(fos);
				fos.close();
			}
			
				catch(FileNotFoundException e)
				{
					e.printStackTrace();
				}
				catch(EncryptedDocumentException e)
				{
					e.printStackTrace();
				}
				catch(IOException e)
				{
					e.printStackTrace();
				}
				
	}

}







