
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class WriteXL {
    String filename = null;
    File file = null;
    FileInputStream fis = null;
    XSSFWorkbook workbook = null;
    XSSFSheet sheet = null ;
    FileOutputStream fos = null ;
    
    public WriteXL(String filename)
    {
        this.filename=filename;
    }
    
    public void modifyCell(int rowNumber,int columnNumber,String sheetName, String valueToWrite) throws IOException{
    	rowNumber -= 1 ;
    	System.out.println("updating sheetName="+ sheetName);
         
    	System.out.println("updating rowNumber="+ Integer.toString(rowNumber));
        System.out.println("updating columnNumber="+ Integer.toString(columnNumber));
        System.out.println("updating valueToWrite="+ valueToWrite);
        
    try {
            file = new File(filename);
            fis = new FileInputStream(file);
            workbook = new XSSFWorkbook(fis);
            sheet = workbook.getSheet(sheetName);
            XSSFRow row = sheet.getRow(rowNumber);
            if (row == null)
            	row = sheet.createRow(rowNumber);
            
            Cell cell = row .getCell(columnNumber);
            
            if (cell == null)
            	cell = row.createCell(columnNumber);
           
            cell.setCellValue(valueToWrite);
            
           
            fos = new FileOutputStream(filename);
            workbook.write(fos);
            
        }catch (Exception e) {
            System.out.println("ERROR : Not able to update the cell");
            e.printStackTrace();
        }
        finally{
        if(fis!= null || fos != null)
        {
        fis.close();
        fos.flush();
        fos.close();
        }
        }
    }
    
    public static void main(String... args) throws IOException
    {
    	List<String> list = Arrays.asList("A", "B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ") ;
    	
    	System.out.println("Exists:"+new File("JavaBooks.xlsx").exists());
        WriteXL writeExcel = new WriteXL("JavaBooks.xlsx");
        writeExcel.modifyCell(13, list.indexOf("G") , "sheet1", "lucas");
        writeExcel.modifyCell(12, list.indexOf("G"), "sheet1", "The lucas");
        
        
    }
}