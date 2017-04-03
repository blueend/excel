package study;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;

public class App {
    private static String getCellValue(CellValue cellValue) {
        String value = "";

        switch (cellValue.getCellType()) {
            case Cell.CELL_TYPE_BOOLEAN:
                value = String.valueOf(cellValue.getBooleanValue());
                break;
            case Cell.CELL_TYPE_NUMERIC:
                value = String.valueOf(cellValue.getNumberValue());
                break;
            case Cell.CELL_TYPE_STRING:
                value = String.valueOf(cellValue.getStringValue());
                break;
            case Cell.CELL_TYPE_BLANK:
                break;
            case Cell.CELL_TYPE_ERROR:
                break;

            // CELL_TYPE_FORMULA will never happen
            case Cell.CELL_TYPE_FORMULA:
                break;
        }

        return value;
    }

    public static void main(String[] args) throws Exception {
        merge(WorkbookFactory.create(new File("data1.xlsx")),
                WorkbookFactory.create(new File("data2.xlsx")));
    }

    public static void merge(Workbook a, Workbook b) throws Exception {
        Workbook mergedWB = new HSSFWorkbook();
        FileOutputStream fileOut = new FileOutputStream("merged.xlsx");

        Sheet newSheet = mergedWB.createSheet();

        int rowIdx = 0;

        FormulaEvaluator aEvaluator = a.getCreationHelper().createFormulaEvaluator();
        FormulaEvaluator bEvaluator = b.getCreationHelper().createFormulaEvaluator();

        rowIdx = sheetToMergedSheet(a, newSheet, rowIdx, aEvaluator);
        rowIdx = sheetToMergedSheet(b, newSheet, rowIdx, bEvaluator);

        mergedWB.write(fileOut);
        fileOut.close();

        System.out.println(String.format("전체 %d 행을 저장하였습니다", rowIdx));
    }

    private static int sheetToMergedSheet(Workbook wb, Sheet newSheet, int rowIdx, FormulaEvaluator evaluator) {
        for (Row row : wb.getSheetAt(0)) {
            Row newRow = newSheet.createRow(rowIdx);

            int cellIdx = 0;
            for (Cell cell : row) {
                Cell newCell = newRow.createCell(cellIdx);

                System.out.println(String.format("[%d,%d] %s", rowIdx, cellIdx, getCellValue(evaluator.evaluate(cell))));
                newCell.setCellValue(getCellValue(evaluator.evaluate(cell)));

                cellIdx++;
            }

            rowIdx++;
        }

        return rowIdx;
    }
}
