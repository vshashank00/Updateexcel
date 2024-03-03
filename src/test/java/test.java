import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class test {
    public static void main(String[] args) throws IOException {


    FileInputStream fileInputStream=new FileInputStream("C://Users//703337600//Downloads//download.xlsx");
    XSSFWorkbook workbook= new XSSFWorkbook(fileInputStream);
        DataFormatter dataFormatter=new DataFormatter();

        int no_of_sheet=workbook.getNumberOfSheets();
    for (int i=0;i<no_of_sheet;i++){
        if (workbook.getSheetName(i).equalsIgnoreCase("sheet1")){
            XSSFSheet sheet=workbook.getSheetAt(i);
            Iterator<Row> row= sheet.iterator();
            int k=0;
            int col=0;
            while (row.hasNext()){
                Row getRow=row.next();
                Iterator<Cell> cell=getRow.iterator();
                while (cell.hasNext()){
                    Cell cell1=cell.next();

                    if (dataFormatter.formatCellValue(cell1).equalsIgnoreCase("price"))
                        System.out.println(col);
                    else col++;
                    if (dataFormatter.formatCellValue(cell1).equalsIgnoreCase("apple")){
                        System.out.println("row:"+k);
                        break;
                    }


                }
                k++;
            }

        }

    }
}
}
