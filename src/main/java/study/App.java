package study;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;


public class App {
    // copy from : http://thinktibits.blogspot.kr/2012/12/POI-iText-Convert-XLS-to-PDF-Java-Program.html
    private static void saveToPDF(String excel, String pdf) throws IOException, DocumentException {
        BaseFont bfKorean = BaseFont.createFont("HYGoThic-Medium", "UniKS-UCS2-H", BaseFont.NOT_EMBEDDED);
        Font fontKorean = new Font(bfKorean, 10, Font.NORMAL);

        FileInputStream input_document = new FileInputStream(new File(excel));
        HSSFWorkbook my_xls_workbook = new HSSFWorkbook(input_document);
        HSSFSheet my_worksheet = my_xls_workbook.getSheetAt(0);
        Iterator<Row> rowIterator = my_worksheet.iterator();
        Document iText_xls_2_pdf = new Document();
        PdfWriter.getInstance(iText_xls_2_pdf, new FileOutputStream(pdf));
        iText_xls_2_pdf.open();
        PdfPTable my_table = new PdfPTable(my_worksheet.getRow(0).getLastCellNum());
        PdfPCell table_cell;

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next(); //Fetch CELL
                switch (cell.getCellType()) { //Identify CELL type
                    case Cell.CELL_TYPE_STRING:
                        table_cell = new PdfPCell(new Phrase(cell.getStringCellValue(), fontKorean));
                        my_table.addCell(table_cell);
                        break;
                }
            }

        }

        iText_xls_2_pdf.add(my_table);
        iText_xls_2_pdf.close();
        input_document.close(); //close xls
    }

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

        saveToPDF("merged.xlsx", "merged.pdf");
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
