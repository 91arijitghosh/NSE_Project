package TestBase;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.time.Duration;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

public class Base_Class
{
    public static Properties prop;
    public static WebDriver driver;

    public Base_Class() {
        try {
            prop = new Properties();
            FileInputStream config = new FileInputStream(System.getProperty("user.dir") + "/src/main/java/Configuration/Config.properties");
            prop.load(config);
        }catch (IOException e)
        {
            e.printStackTrace();
        }
    }

    public static void Initialize()
    {
        System.setProperty("webdriver.chrome.driver",System.getProperty("user.dir")+prop.getProperty("Driver_Location"));
        driver = new ChromeDriver();
        driver.manage().window().maximize();
        //driver.manage().deleteAllCookies();
        driver.get(prop.getProperty("URL"));
        //driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(10));
        //driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
       // driver.navigate().refresh();

    }

}
