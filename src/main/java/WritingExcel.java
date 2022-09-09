import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class WritingExcel {
    public static void main(String[] args) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Data Information");


        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[]{"NAME", "LAST NAME", "EMAIL","PASSWORD","COMPANY","ADDRESS","CITY","ZIP_CODE","MOBILE_PHONE"});
        data.put("2", new Object[]{"LIONEL", "MESSI", "LIONELMESSI@HOTMAIL.COM","MESIAS10","PSG","ARGENTINA","PARIS","23456","556789321"});
        data.put("3", new Object[]{"CRISTIANO", "RONALDO", "CR7@HOTMAIL.COM","ELBICHOO","UNITED","PORTUGAL","ENGLAND","23457","556789331"});


        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset) {
            Row row = sheet.createRow(rownum++);
            Object[] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof String)
                    cell.setCellValue((String) obj);
                else if (obj instanceof Integer)
                    cell.setCellValue((Integer) obj);
            }
        }
        try {
            FileOutputStream out = new FileOutputStream(new File("src/main/resources/DataInformation.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("DataInformation.xlsx written successfully on disk.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
