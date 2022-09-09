import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReadingIteratorExcel {

    public static void main(String[] args) throws IOException {

        String excelFilePath="src/main/resources/JavaFile.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);

        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);

        //Iterator
        Iterator iterator = sheet.iterator();//Return all rows

        while(iterator.hasNext()){
            XSSFRow row = (XSSFRow) iterator.next();
            Iterator cellIterator = row.cellIterator();

            while (cellIterator.hasNext()){
                XSSFCell cell = (XSSFCell)cellIterator.next();
                switch(cell.getCellType()){
                    case STRING:System.out.print(cell.getStringCellValue());
                        break;
                    case NUMERIC:System.out.print(cell.getNumericCellValue());
                        break;
                    case BOOLEAN:System.out.print(cell.getBooleanCellValue());
                        break;
                }
                System.out.print(" | ");
            }
            System.out.println();
        }

    }


}
