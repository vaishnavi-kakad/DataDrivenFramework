package ExcelUtilsp;

public class ExcelSheet {
public static void main(String[] args) {
	//String projectPath= System.getProperty("user.dir");
	ExcelFileOp excel= new ExcelFileOp("C:\\Users\\Bhujang Kakad\\Documents\\SimpleProject\\excel\\travel.xlsx","Sheet1");
			excel.getRowCount();
			excel.getCellData(0, 0);
			excel.getCellDataNumber(1, 1);
}

}
