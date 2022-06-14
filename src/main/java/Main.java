import java.util.List;

public class Main {

    public static void main(String[] args) {

        ExcelOperations boj = new ExcelOperations();
        boj.readSheetByRange("Supplier!A1:H");
    }
}
