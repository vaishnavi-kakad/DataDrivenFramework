package ExcelUtilsp;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileOp {
	
	//static String projectPath;
	static XSSFWorkbook wb;
	static XSSFSheet sheet;
	
	
	public ExcelFileOp(String excelpath, String sheetName) {
		try {
		
		 wb = new XSSFWorkbook( excelpath);
		 sheet = wb.getSheet(sheetName);}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}
	public static void main(String[] args) {
		getRowCount();
		getCellData(0,0);
		getCellDataNumber(1,1);
		getcolCount();
	}

	public static int getRowCount() {
		int rowCount =0;
		try {
			
		 rowCount = sheet.getPhysicalNumberOfRows();
		System.out.println("no of rows"+rowCount);
	} catch (Exception exp) {
		
		System.out.println(exp.getMessage());
		System.out.println(exp.getCause());
		exp.printStackTrace();
	}
		return rowCount;
	}
	
	public static String getCellData(int rowNum, int colNum) {
		String cellData= null;
		try {
	
		  cellData = sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
		 // System.out.println(cellData);
		} catch (Exception exp) {
			
			System.out.println(exp.getMessage());
			System.out.println(exp.getCause());
			exp.printStackTrace();
		}
		return cellData;
	}
	public static int getcolCount() {
		int colCount =0;
		try {
			
		 colCount = sheet.getRow(0).getPhysicalNumberOfCells();
		System.out.println("no of col"+colCount);
	} catch (Exception exp) {
		
		System.out.println(exp.getMessage());
		System.out.println(exp.getCause());
		exp.printStackTrace();
	}
		return colCount;
	}
	
	public static void getCellDataNumber(int rowNum, int colNum) {
		try {
		  String cellData= sheet.getRow(1).getCell(1).getStringCellValue();
		 // System.out.println(cellData);
		} catch (Exception exp) {
			
			System.out.println(exp.getMessage());
			System.out.println(exp.getCause());
			exp.printStackTrace();
		}
		
	}
}