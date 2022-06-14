import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;

public class CommonUtility {

    public static Object getCellValueAsObject(Cell value) {

        Object cellValue = null;
        try {
            switch (value.getCellType()) {
                case BOOLEAN:
                    cellValue = value.getBooleanCellValue();
                    break;
                case NUMERIC:
                    cellValue = value.getNumericCellValue();
                    break;
                case STRING:
                    cellValue = value.getStringCellValue();
                    break;
                case BLANK:
                    cellValue = "";
                    break;
                case ERROR:
                    cellValue = value.getErrorCellValue();
                    break;
                case FORMULA:
                    cellValue = ((XSSFCell) value).getRawValue();
                    break;
            }
        } catch (NullPointerException npex) {
            System.out.println(npex.getMessage());
            cellValue = "";
        }

        return cellValue;
    }
}
