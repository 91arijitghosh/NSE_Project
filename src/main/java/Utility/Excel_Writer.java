package Utility;

import org.apache.poi.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class Excel_Writer
{
    public XSSFWorkbook book;
    public XSSFSheet sheet;
    public XSSFWorkbook input_book;
    public XSSFSheet input_sheet;
    public XSSFRow row;
    public XSSFCell cell;
    public double Calls_OI;
    public double Puts_OI;

    public Excel_Writer()
    {
        Calls_OI = 0;
        Puts_OI = 0;
    }


    public void Write_Data_to_Excel(XSSFWorkbook book1) throws IOException {

        File fl = new File(System.getProperty("user.dir")+"/src/main/java/Output/NSE.xlsx");
        FileOutputStream fos = new FileOutputStream(fl);
        book1.write(fos);
        fos.close();
        book1.close();
    }
    public void Read_Data_from_Excel() throws IOException {

        File fl = new File(System.getProperty("user.dir")+"/src/main/java/Output/NSE.xlsx");
        FileInputStream fis = new FileInputStream(fl);
        input_book = new XSSFWorkbook(fis);
        input_sheet = input_book.getSheetAt(0);
        int row_nums = input_sheet.getLastRowNum();
        int call_nums = input_sheet.getRow(0).getLastCellNum();

        XSSFRow Row = input_sheet.getRow(row_nums);
        XSSFCell Cell1 = Row.getCell(0);
        XSSFCell Cell2 = Row.getCell(call_nums-1);

        Calls_OI = Excel_Formula_Evaluator(input_book,Cell1);
        Puts_OI = Excel_Formula_Evaluator(input_book,Cell2);

    }

    public void Write_Data_to_Excel() throws IOException {
        File fl = new File(System.getProperty("user.dir")+"/src/main/java/Output/Difference.xlsx");
        FileInputStream fis = new FileInputStream(fl);
        book = new XSSFWorkbook(fis);
        sheet = book.getSheetAt(0);
        int ronum = sheet.getLastRowNum()+1;
        row = sheet.createRow(ronum);
        Set_Date(row,0);
        SetCell(row,1,Calls_OI);
        SetCell(row,2,Puts_OI);
        SetCellDiff(row,3,ronum);
        SetCellRatio(row,4,ronum);

        FileOutputStream fos = new FileOutputStream(fl);
        book.write(fos);
        fis.close();
        book.close();
        fos.close();

    }
    public void Set_Date(XSSFRow row,int cellIndex)
    {
        XSSFCell cell = row.createCell(cellIndex);
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
        LocalDateTime now = LocalDateTime.now();
        cell.setCellValue(dtf.format(now));

    }

    public void SetCell(XSSFRow row,int cellIndex, String Value, CellStyle style)
    {
        XSSFCell cell = row.createCell(cellIndex);
        cell.setCellValue(Value);
        cell.setCellStyle(style);

    }

    public void SetCell(XSSFRow row,int cellIndex, Double Value)
    {
        XSSFCell cell = row.createCell(cellIndex);
        cell.setCellValue(Value);

    }
    public void SetCellDiff(XSSFRow row,int cellIndex,int rownum)
    {
        String n = Integer.toString(rownum+1);
        XSSFCell cell = row.createCell(cellIndex);
        cell.setCellFormula("B"+n+"-C"+n);

    }
    public void SetCellRatio(XSSFRow row,int cellIndex,int rownum)
    {
        String n = Integer.toString(rownum+1);
        XSSFCell cell = row.createCell(cellIndex);
        cell.setCellFormula("C"+n+"/B"+n);

        CellStyle style = book.createCellStyle();
        Double indicate = 0.0;
        indicate = Excel_Formula_Evaluator(book,cell);
        if(indicate>1.0)
        {
            XSSFCell cella = row.createCell(cellIndex+1);
            cella.setCellValue("Buy");
            style.setFillForegroundColor(IndexedColors.GREEN.index);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cella.setCellStyle(style);

        }else if(indicate<1.0)
        {
            XSSFCell cella = row.createCell(cellIndex+1);
            cella.setCellValue("Sell");
            style.setFillForegroundColor(IndexedColors.RED.index);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cella.setCellStyle(style);

        }else
        {
            XSSFCell cella = row.createCell(cellIndex+1);
            cella.setCellValue("Nutral");
            style.setFillForegroundColor(IndexedColors.BLUE.index);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cella.setCellStyle(style);
        }

    }

    public double Excel_Formula_Evaluator(XSSFWorkbook bk, XSSFCell cello)
    {
        double val = 0;
        FormulaEvaluator evaluator = bk.getCreationHelper().createFormulaEvaluator();

        if (cello.getCellType() == CellType.FORMULA )
        {
            switch (evaluator.evaluateFormulaCell(cello))
            {
                case BOOLEAN:
                    System.out.println(cello.getBooleanCellValue());
                    break;
                case NUMERIC:
                    val = cello.getNumericCellValue();
                    System.out.println(cello.getNumericCellValue());
                    break;
                case STRING:
                    val = cello.getNumericCellValue();
                    System.out.println(cello.getStringCellValue());
                    break;
                case FORMULA:
                    val = cello.getNumericCellValue();
                    System.out.println(cello.getNumericCellValue());
                    break;
            }
        }
        return val;
    }

}
