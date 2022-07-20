import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class ExcelDataDriven {

    public ArrayList<String> getDataFromExcel(String testcaseName) throws IOException {
        // file input stream argument
        FileInputStream fis = new FileInputStream("src/main/resources/data.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        ArrayList<String> arrayList = new ArrayList<>();

        int sheets = workbook.getNumberOfSheets();
        for (int i = 0; i < sheets; i++) {
            if (workbook.getSheetName(i).equalsIgnoreCase("testdata")){
                XSSFSheet sheet = workbook.getSheetAt(i);

                // Identify Testcases column by scaning the entire 1st row
                Iterator<Row> rows = sheet.iterator();  // sheet is collection of rows
                Row firstrow = rows.next();
                Iterator<Cell> cells = firstrow.cellIterator(); //row is collection of cells

                int k = 0;
                int column = 0;
                while (cells.hasNext()){
                    Cell cell = cells.next();
                    if (cell.getStringCellValue().equalsIgnoreCase("TestCases")){
                        //desired column
                        column = k;
                    }
                    k++;
                }
                // once column is identified then scan entire selected column to identify one cell
                while (rows.hasNext()){
                    Row row = rows.next();
                    if (row.getCell(column).getStringCellValue().equalsIgnoreCase(testcaseName)) {
                        // after you grab purchase testcase row pull all the data of that row into test
                        Iterator<Cell> cv = row.cellIterator();
                        while (cv.hasNext()) {
                            Cell cell = cv.next();
                            if (cell.getCellTypeEnum() == CellType.STRING){
                                arrayList.add(cell.getStringCellValue());
                            } else {
                               arrayList.add(String.valueOf(cell.getNumericCellValue()));
                            }
                        }
                    }
                }
            }
        }
        return arrayList;
    }
}
