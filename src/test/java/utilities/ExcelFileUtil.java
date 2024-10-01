package utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileUtil {
	Workbook wb;
	//constructor for reading excel path
	public ExcelFileUtil(String Excelpath)throws Throwable
	{
		FileInputStream fi  = new FileInputStream(Excelpath);
		wb = WorkbookFactory.create(fi);
	}
	//method for counting no of rows in a sheet
	public int rowCount(String SheetName)
	{
		return wb.getSheet(SheetName).getLastRowNum();
	}
	//method for reading cell data
	public String getCellData(String sheetname,int row,int column)
	{
		DataFormatter df = new DataFormatter();
		String data = df.formatCellValue(wb.getSheet(sheetname).getRow(row).getCell(column));
		/*
	    String data="";
		if(wb.getSheet(sheetname).getRow(row).getCell(column).getCellType()==CellType.NUMERIC)
		{
			int celldata = (int)wb.getSheet(sheetname).getRow(row).getCell(column).getNumericCellValue();
			data =String.valueOf(celldata);
		}
		else
		{
			data = wb.getSheet(sheetname).getRow(row).getCell(column).getStringCellValue();
		}*/
		return data;
	}
	//method for writing cell data
	public void setCellData(String sheetName,int row,int column,String status,String WriteExcel)throws Throwable
	{
		//get sheet from wb
		Sheet ws = wb.getSheet(sheetName);
		//get row from sheet
		Row rowNum = ws.getRow(row);
		//create cell
		Cell cell = rowNum.createCell(column);
		//write status
		cell.setCellValue(status);
		if(status.equalsIgnoreCase("Pass"))
		{
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			font.setColor(IndexedColors.GREEN.getIndex());
			font.setBold(true);
			style.setFont(font);
			ws.getRow(row).getCell(column).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("Fail"))
		{
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			font.setColor(IndexedColors.RED.getIndex());
			font.setBold(true);
			style.setFont(font);
			ws.getRow(row).getCell(column).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("Blocked"))
		{
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			font.setColor(IndexedColors.BLUE.getIndex());
			font.setBold(true);
			style.setFont(font);
			ws.getRow(row).getCell(column).setCellStyle(style);
		}
		FileOutputStream fo = new FileOutputStream(WriteExcel);
		wb.write(fo);
	}

	public static void main(String[] args) throws Throwable {

		ExcelFileUtil xl = new ExcelFileUtil("D:\\Selenium Folder\\LoginData.xlsx");

		//count no of rows in sheet

		int rc =xl.rowCount("LoginData");

		System.out.println(rc);

		for(int i=1;i<=rc;i++)
		{
			String username =xl.getCellData("LoginData", 1, 0);
			String password = xl.getCellData("LoginData", 1, 1);
			String phoneno = xl.getCellData("LoginData", 1, 2);
			System.out.println(username+" "+password+" "+phoneno);

			//xl.setCellData("LoginData", i, 3, "Pass", "D:/Results.xlsx");
			 xl.setCellData("LoginData", i, 3, "Fail", "D:/Results.xlsx");
			//xl.setCellData("LoginData", i, 3, "Blocked", "D:/Results.xlsx");
		}
	}
}