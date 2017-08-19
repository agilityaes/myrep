package yahoo;

import java.util.List;

import java.util.concurrent.TimeUnit;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.testng.Reporter;
import org.testng.annotations.Listeners;

import atu.testng.reports.ATUReports;
import atu.testng.reports.listeners.ATUReportsListener;
import atu.testng.reports.listeners.ConfigurationListener;
import atu.testng.reports.listeners.MethodListener;
import atu.testng.reports.logging.LogAs;
import atu.testng.selenium.reports.CaptureScreen;
import atu.testng.selenium.reports.CaptureScreen.ScreenshotOf;


@Listeners({ConfigurationListener.class,MethodListener.class,ATUReportsListener.class})
public class Home extends DriverClass
{
  
	{
		System.setProperty("atu.reporter.config", "e:\\agility\\atu.properties");
	}
  public void open()
  {
	driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
	driver.manage().window().maximize();
	driver.get("http://www.yahoomail.com");  
  }
  public void createacc() throws Exception
  {
	open();
	Thread.sleep(3000);
	driver.findElement(By.id("createacc")).click();
	Thread.sleep(5000);
	System.out.println("updated by geetha");
	
	try
	{
		if(driver.findElement(By.name("firstName")).isDisplayed())
		{
			ATUReports.add("Reg page is displayed", LogAs.PASSED,new CaptureScreen(ScreenshotOf.DESKTOP));
			driver.findElement(By.name("firstName")).sendKeys("abcd");
		}
	}
	catch(Exception obj)
	{
		Reporter.log("<font color='red'><b>Registration page not displayed</b></font>");
	}
	
	
	
	//enter lastname, yahooid, select month,day, year....etc
	
  }
  public void login() throws Exception
  {
	open();
	driver.findElement(By.name("username")).sendKeys("venkat123456a");
	driver.findElement(By.name("signin")).click();
	Thread.sleep(5000);		
	driver.findElement(By.id("login-passwd")).sendKeys("mqq987654");
	driver.findElement(By.name("verifyPassword")).click();
	Thread.sleep(5000);
  }
}
