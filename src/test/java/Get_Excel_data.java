import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class Get_Excel_data {
    public static void main(String[] args) throws IOException {
        FileInputStream fileInputStream=new FileInputStream("C:/Users/703337600/Desktop/Book1.xlsx");
        XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream);
        ArrayList<String> arrayList=new ArrayList<>();
        int sheets=workbook.getNumberOfSheets();
        for (int i=0;i<sheets;i++){
            if (workbook.getSheetName(i).equalsIgnoreCase("Sheet1")){
                XSSFSheet sheet=workbook.getSheetAt(i);
                Iterator <Row> row= sheet.iterator();
                Row firstrow=row.next();
               Iterator<Cell> cell= firstrow.cellIterator();
               int k=0;
               int col=0;
               while (cell.hasNext()){
                      Cell value= cell.next();

                     if( value.getStringCellValue().equalsIgnoreCase("Testcase"))
                   {
                      col=k;
                   }
                   k++;
               }
                System.out.println(col);
               while (row.hasNext()){
                   Row r= row.next();
                   if(r.getCell(col).getStringCellValue().equalsIgnoreCase("purchase")){
                      Iterator<Cell>cb=r.cellIterator();
                      while (cb.hasNext()) {
                          Cell c = cb.next();
                          if (c.getCellType() == CellType.STRING)
                              arrayList.add(c.getStringCellValue());
                          else
                          {
                              arrayList.add(NumberToTextConverter.toText(c.getNumericCellValue()));}
                      }
                       System.out.println(arrayList);
                   }

               }

            }
        }

    }
}
//vedio 227 from rahul shetty course