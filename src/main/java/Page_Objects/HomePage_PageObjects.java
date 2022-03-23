package Page_Objects;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;

import java.util.List;

public class HomePage_PageObjects
{
    @FindBy(xpath = "(//th[@title='Open Interest in contracts'])[1]")
    public WebElement CALLS_OI;

    @FindBy(xpath = "(//th[@title='Change in Open Interest (Contracts)'])[1]")
    public WebElement CALLS_CHNG_IN_OI;

    @FindBy(xpath = "(//th[@title='Last Traded Price'])[1]")
    public WebElement CALLS_LTP;

    @FindBy(xpath = "//th[text()='Strike Price']")
    public WebElement CALLS_STRIKE_PRICE;

    @FindBy(xpath = "//table[1]//tbody//tr")
    public List<WebElement> table_length;


    //
    //thead//tr[2]//th[22]



    //
}
