package ExcelUtilsp;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class ExcelDataProvider{
	WebDriver driver = null;
	@BeforeTest
	public void setUpTest() 
	{
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\Bhujang Kakad\\Documents\\SimpleProject\\driver\\chromedriver\\chromedriver.exe");
		driver = new ChromeDriver();
	}
	


	@Test(dataProvider ="test1data")
	public void test1(String username, String password) throws InterruptedException
	{
		driver.get("https://github.com/login");
		driver.findElement(By.name("login")).sendKeys(username);
		driver.findElement(By.name("password")).sendKeys(password);
		driver.findElement(By.name("commit")).click();
		Thread.sleep(2000);
	}
	
	@DataProvider(name="test1data")
	public Object[][] getData()
	{
		String excelPath= "C:\\Users\\Bhujang Kakad\\Documents\\SimpleProject\\excel\\travel.xlsx"; 
		Object data[][] =testData(excelPath,"Sheet1");
		return data;
	}
	
	
	public Object[][] testData(String excelpath, String sheetName) {
	ExcelFileOp excel= new ExcelFileOp("C:\\Users\\Bhujang Kakad\\Documents\\SimpleProject\\excel\\travel.xlsx","Sheet1");
	
	int rowCount = excel.getRowCount();
	int colCount = excel.getcolCount();
	Object data[][]= new Object[rowCount-1][colCount];
	for(int i=1; i<rowCount;i++) {
		for(int j=0;j<colCount;j++)
		{
			String cellData= excel.getCellData(i, j);
			//System.out.print(cellData+ "|");
			data[i-1][j]= cellData;
		}
		//System.out.println();
	}
	return data;
}
}