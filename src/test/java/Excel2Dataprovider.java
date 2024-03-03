import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFRangeCopier;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.Type;
import java.util.Arrays;

public class Excel2Dataprovider {
    @Test(dataProvider = "data")
    public void exceltoData( String datano,String login,String purchase,String addprofile) throws IOException {
        System.out.println(datano+login+purchase+addprofile);


    }
    @DataProvider( name = "data")
    Object[][] grabdata() throws IOException {
        FileInputStream fileInputStream=new FileInputStream("C:/Users/703337600/Desktop/Book1.xlsx");
        XSSFWorkbook xssfWorkbook=new XSSFWorkbook(fileInputStream);
        DataFormatter dataFormatter=new DataFormatter();

        int i= xssfWorkbook.getNumberOfSheets();
        for (int j=0;j<i;j++){
            if (xssfWorkbook.getSheetName(j).equalsIgnoreCase("sheet1")){
                XSSFSheet sheet=xssfWorkbook.getSheetAt(j);
                int nrow=sheet.getPhysicalNumberOfRows();
                XSSFRow row=sheet.getRow(0);
                int col=row.getLastCellNum();
                Object data1[][]=new Object[nrow-1][col];//we have nrow-1 because we don want first row that is heading fof col
                for (int k=0;k<nrow-1;k++){
                    row=sheet.getRow(k+1);
                    for (int s=0;s<col;s++){

                        data1[k][s]=dataFormatter.formatCellValue(row.getCell(s));
                        System.out.println(data1[k][s]);
//                        if (row.getCell(s).getCellType() == CellType.STRING)
//                             data[k][s]=row.getCell(s).getStringCellValue();
//                        else
//                            data[k][s]= NumberToTextConverter.toText(row.getCell(s).getNumericCellValue());
//                        System.out.println(data[k][s]);

                    }}


                System.out.println(Arrays.deepToString(data1));
return data1;
        }}


        return new Object[0][0];
    }

}
