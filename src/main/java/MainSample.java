import java.io.IOException;
import java.util.ArrayList;

public class MainSample {

    public static void main(String[] args) throws IOException {

        ExcelDataDriven excelDataDriven = new ExcelDataDriven();
        ArrayList data = excelDataDriven.getDataFromExcel("Add Profile");
        System.out.println(data.get(0));
        System.out.println(data.get(1));
        System.out.println(data.get(2));
        System.out.println(data.get(3));
    }
}
