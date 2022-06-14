import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelOperations {

    public List<List<String>> readSheetByRange(String sheetRange) {
        List<List<String>> sheetAsList = new ArrayList<>();

        try {
            String sheetName = sheetRange.split("!")[0];
            String gridRange = sheetRange.split("!")[1];

            FileInputStream file = new FileInputStream("mde-2021-22.xlsx");
            int startColIndex = 0, startRowIndex = 0, lastColIndex = 0;

            try {
                startColIndex = CellReference.convertColStringToIndex(gridRange.split(":")[0].replaceAll("\\d", ""));
                startRowIndex = Integer.parseInt(gridRange.split(":")[0].replaceAll("[a-zA-Z]", ""));
                lastColIndex = CellReference.convertColStringToIndex(sheetRange.split("!")[1].split(":")[1]);
            } catch (NumberFormatException nfex) {
                System.out.println("Passed range for reading the sheet is not correct, Please note the format" +
                        " is 'SheetName!StartCellIndex:EndCol'");
                System.exit(0);
            }

            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheet(sheetName);

            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            //Iterate through each row one by one
            for (int j = 0; j <= sheet.getLastRowNum(); j++) {

                if (j+1 < startRowIndex)
                    continue;

                List<String> rowAsList = new ArrayList<>();
                Row row = sheet.getRow(j);

                if (row != null) {
                    //Iterate through each cell one by one
                    for (int i = 0; i <= lastColIndex; i++) {
                        if (i < startColIndex)
                            continue;
                        try {
                            if (row.getCell(i).getCellType() == CellType.FORMULA)
                                rowAsList.add(((XSSFCell) row.getCell(i)).getRawValue());
                            else
                                rowAsList.add(formatter.formatCellValue(row.getCell(i)));
                        } catch (NullPointerException npex) {
                            System.out.println(npex.getMessage());
                            rowAsList.add("");
                        }
                    }
                }
                sheetAsList.add(rowAsList);
            }
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return sheetAsList;
    }
}
