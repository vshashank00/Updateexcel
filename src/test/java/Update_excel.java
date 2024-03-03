import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Iterator;

public class Update_excel {
    String filename="C://Users//703337600//Downloads//download.xlsx";
    @Test
    void dowload() throws InterruptedException, IOException {
        WebDriverManager.chromedriver().setup();
        WebDriver driver= new ChromeDriver();
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(7));
        driver.get("https://rahulshettyacademy.com/upload-download-test/index.html");
        driver.findElement(By.xpath("//button[@id=\"downloadButton\"]")).click();
        WebElement element=driver.findElement(By.xpath("//input[@type='file']"));
        Thread.sleep(2000);
        int row=getrow();
        int col=getcol();
        update(row,col,253);
        element.sendKeys("C:/Users/703337600/Downloads/download.xlsx");
        WebDriverWait wait=new WebDriverWait(driver,Duration.ofSeconds(5));
        wait.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath("//div[contains(text(),\"Updated\")]"))));
        String text=driver.findElement(By.xpath("//div[contains(text(),'Updated')]")).getText();
        Assert.assertEquals("Updated Excel Data Successfully.",text);
        wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//div[contains(text(),\"Updated\")]")));
        String no=driver.findElement(By.xpath("//div[text()='Price']")).getAttribute("data-column-id");
        System.out.println(driver.findElement(By.xpath("//div[text()='Apple']/parent::div/following-sibling::div[@data-column-id="+no+"]/child::div")).getText());
driver.quit();
}
void update( int row,int col,int value) throws IOException {
    FileInputStream fileInputStream=new FileInputStream(filename);
    XSSFWorkbook workbook= new XSSFWorkbook(fileInputStream);
    XSSFSheet sheet=workbook.getSheetAt(0);
    Row row1=sheet.getRow(row);
    Cell cell= row1.getCell(col);
    cell.setCellValue(value);
    FileOutputStream fileOutputStream=new FileOutputStream(filename);
    workbook.write(fileOutputStream);//it will save the chages in file


}
int getrow() throws IOException {

    FileInputStream fileInputStream=new FileInputStream(filename);
    XSSFWorkbook workbook= new XSSFWorkbook(fileInputStream);
    DataFormatter dataFormatter=new DataFormatter();

    int no_of_sheet=workbook.getNumberOfSheets();
    for (int i=0;i<no_of_sheet;i++){
        if (workbook.getSheetName(i).equalsIgnoreCase("sheet1")){
            XSSFSheet sheet=workbook.getSheetAt(i);
            Iterator<Row> row= sheet.iterator();
            int k=0;
            while (row.hasNext()){
                Row getRow=row.next();
                Iterator<Cell> cell=getRow.iterator();
                while (cell.hasNext()){
                    Cell cell1=cell.next();
                    if (dataFormatter.formatCellValue(cell1).equalsIgnoreCase("apple")){
                        System.out.println("row:"+k);
                        return k;
                    }


                }
                k++;
            }

        }
    }
    return -1;
}
int getcol() throws IOException {
        FileInputStream fileInputStream= new FileInputStream(filename);
    DataFormatter dataFormatter=new DataFormatter();
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet=xssfWorkbook.getSheet("Sheet1");
        Row row=sheet.getRow(0);
        Iterator <Cell>getcol=row.cellIterator();
        int col=0;
        while (getcol.hasNext()){
            Cell cell=getcol.next();
            if(dataFormatter.formatCellValue(cell).equalsIgnoreCase("price")){
                return col;
            }else col++;

        }

    return -1;
}
}
