package yahoo;

import java.io.FileInputStream;
import java.lang.reflect.Method;
import java.net.URL;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Platform;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.testng.annotations.Listeners;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import atu.testng.reports.listeners.ATUReportsListener;
import atu.testng.reports.listeners.ConfigurationListener;
import atu.testng.reports.listeners.MethodListener;


@Listeners({ConfigurationListener.class,MethodListener.class,ATUReportsListener.class})
public class ReTest  extends DriverClass
{
  	
	//DesiredCapabilities ds;
	{
		System.setProperty("atu.reporter.config", "e:\\agility\\atu.properties");
	}
  @Test
  @Parameters({"browser"})
  public void retesting(String br) throws Exception
  {
	  if(br.matches("firefox"))
	  {
		 driver=new FirefoxDriver();		 
		  //ds=DesiredCapabilities.firefox();
		  //ds.setPlatform(Platform.WINDOWS);
	  }
	  if(br.matches("ie"))
	  {
		  System.setProperty("webdriver.ie.driver","c:\\IEDriverServer.exe");
		  driver=new InternetExplorerDriver();	
		  //ds=DesiredCapabilities.internetExplorer();
		  //ds.setPlatform(Platform.WINDOWS);
	  }
	  //driver=new RemoteWebDriver(new URL("http://10.138.75.144:1234/wd/hub"), ds);
	  String classname,methodname;
	  FileInputStream fin =new FileInputStream("e:\\agility\\testdata.xlsx"); //excel file for reading
	  XSSFWorkbook wb=new XSSFWorkbook(fin);        // get workbook in the excel file
	  XSSFSheet ws=wb.getSheet("retest"); //get sheet in the workbook
	  Row row;
	  for(int r=1;r<=ws.getLastRowNum();r++)  //for all the rows in the sheet
	  {
		row=ws.getRow(r);
		if(row.getCell(4).getStringCellValue().matches("yes"))
		{
			classname=row.getCell(2).getStringCellValue();
			methodname=row.getCell(3).getStringCellValue();
			Class c=Class.forName(classname);  
			Method m=c.getMethod(methodname, null); 
			Object obj=c.newInstance();   
			m.invoke(obj, null); 			
		}
	  }
	  fin.close();
  }
}













