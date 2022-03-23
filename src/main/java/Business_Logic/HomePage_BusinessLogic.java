package Business_Logic;

import Page_Objects.HomePage_PageObjects;
import TestBase.Base_Class;
import Utility.Excel_Writer;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.pagefactory.AjaxElementLocatorFactory;

import java.io.IOException;
import java.util.List;

public class HomePage_BusinessLogic extends Base_Class
{
    HomePage_PageObjects home = null;
    XSSFWorkbook book;
    XSSFSheet sheet;
    Excel_Writer xl;
    CellStyle style1;
    CellStyle style2;
    XSSFRow row;
    XSSFCell cell;
    int j = 0;
    String col_letter = "";
    double Call_OI_SUM=0;
    double Put_OI_SUM=0;
    public HomePage_BusinessLogic()
    {
        home = new HomePage_PageObjects();
        AjaxElementLocatorFactory factory = new AjaxElementLocatorFactory(driver,10);
        PageFactory.initElements(factory,home);
    }

    public void iterate_data(int input1,int input2) throws IOException {

          book = new XSSFWorkbook();
          sheet = book.createSheet("NSE_Data");
          xl = new Excel_Writer();
          style1 = book.createCellStyle();
          style1.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
          style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);

          style2 = book.createCellStyle();
          style2.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE1.getIndex());
          style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        //System.out.println(home.CALLS_OI.getText()+" : "+home.CALLS_CHNG_IN_OI.getText()+" : "+home.CALLS_LTP+" : "+home.CALLS_STRIKE_PRICE);
        row = sheet.createRow(0);
        xl.SetCell(row,0,"CALLS_OI",style1);
        xl.SetCell(row,1,"CALLS_CHNG_IN_OI",style1);
        xl.SetCell(row,2,"CALLS_LTP",style1);
        xl.SetCell(row,3,"STRIKE_PRICE",style1);
        xl.SetCell(row,4,"PUTS_LTP",style1);
        xl.SetCell(row,5,"PUTS_CHNG_IN_OI",style1);
        xl.SetCell(row,6,"PUTS_OI",style1);



        for (int i =1;i<=home.table_length.size();i++)
        {
            String CALLS_OI = driver.findElement(By.xpath("//table[1]//tbody//tr["+i+"]//td[2]")).getText().replace(",","");
            String CALLS_CHNG_IN_OI = driver.findElement(By.xpath("//table[1]//tbody//tr["+i+"]//td[3]")).getText().replace(",","");
            String CALLS_LTP = driver.findElement(By.xpath("//table[1]//tbody//tr["+i+"]//td[6]")).getText().replace(",","");
            String STRIKE_PRICE = driver.findElement(By.xpath("//table[1]//tbody//tr["+i+"]//td[12]")).getText().replace(",","");
            STRIKE_PRICE = STRIKE_PRICE.replace(",","");
            STRIKE_PRICE = STRIKE_PRICE.replace(".00","");
            String PUTS_LTP = driver.findElement(By.xpath("//table[1]//tbody//tr["+i+"]//td[18]")).getText().replace(",","");
            String PUTS_CHNG_IN_OI = driver.findElement(By.xpath("//table[1]//tbody//tr["+i+"]//td[21]")).getText().replace(",","");
            String PUTS_OI = driver.findElement(By.xpath("//table[1]//tbody//tr["+i+"]//td[22]")).getText().replace(",","");

            if(Integer.parseInt(STRIKE_PRICE) >= input1 && Integer.parseInt(STRIKE_PRICE) <= input2)
            {
                j++;
                sheet.createRow(j);
                sheet.getRow(j).createCell(0).setCellValue(Double.parseDouble(CALLS_OI));
                sheet.getRow(j).createCell(1).setCellValue(Double.parseDouble(CALLS_CHNG_IN_OI));
                sheet.getRow(j).createCell(2).setCellValue(Double.parseDouble(CALLS_LTP));
                sheet.getRow(j).createCell(3).setCellValue(Double.parseDouble(STRIKE_PRICE));
                sheet.getRow(j).createCell(4).setCellValue(Double.parseDouble(PUTS_LTP));
                sheet.getRow(j).createCell(5).setCellValue(Double.parseDouble(PUTS_CHNG_IN_OI));
                sheet.getRow(j).createCell(6).setCellValue(Double.parseDouble(PUTS_OI));
                //System.out.println(OI + " : " + CHNG_IN_OI + " : " + LTP + " : " + STRIKE_PRICE);
            }
        }
        j++;
        sheet.createRow(j);
        cell = sheet.getRow(j).createCell(0);
        col_letter = CellReference.convertNumToColString(cell.getColumnIndex());
        cell.setCellFormula("SUM("+col_letter+"2:"+col_letter+j+")");
        cell.setCellStyle(style2);

        cell = sheet.getRow(j).createCell(6);
        col_letter = CellReference.convertNumToColString(cell.getColumnIndex());
        cell.setCellFormula("SUM("+col_letter+"2:"+col_letter+j+")");
        cell.setCellStyle(style2);

        xl.Write_Data_to_Excel(book);
        book.close();
        xl.Read_Data_from_Excel();
        xl.Write_Data_to_Excel();


    }


}
