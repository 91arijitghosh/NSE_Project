import Business_Logic.HomePage_BusinessLogic;
import TestBase.Base_Class;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import java.io.IOException;

public class Get_NSE_DATA extends Base_Class
{
    @Parameters({"Value1","Value2"})
    @Test
    public void Get_Data( Integer Value1, Integer Value2) throws InterruptedException, IOException {
        Initialize();
        HomePage_BusinessLogic home = new HomePage_BusinessLogic();
        Thread.sleep(2000);
        home.iterate_data(Value1,Value2);

    }

    @AfterMethod
    public void quit_browser()
    {
        driver.quit();
    }
}
