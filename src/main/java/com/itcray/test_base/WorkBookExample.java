package com.itcray.test_base;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;

import javax.print.DocFlavor.STRING;

// import com.spire.pdf.tables.table.DataTable;
import com.spire.xls.*;


// import com.itextpdf.text.Document;
// import com.itextpdf.text.Paragraph;
// import com.itextpdf.text.Phrase;
// import com.itextpdf.text.Rectangle;
// import com.itextpdf.text.pdf.PdfPCell;
// import com.itextpdf.text.pdf.PdfPTable;
// import com.itextpdf.text.pdf.PdfWriter;

// import org.apache.poi.ss.usermodel.BuiltinFormats;
// import org.apache.poi.ss.usermodel.Cell;
// import org.apache.poi.ss.usermodel.CellStyle;
// import org.apache.poi.ss.usermodel.CellType;
// import org.apache.poi.ss.usermodel.DataFormatter;
// import org.apache.poi.ss.usermodel.Row;
// import org.apache.poi.ss.usermodel.Sheet;
// import org.apache.poi.ss.usermodel.Workbook;
// import org.apache.poi.ss.usermodel.WorkbookFactory;
// import org.apache.poi.xssf.usermodel.XSSFSheet;
// import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkBookExample {
    /**
     * 原始Excel模板地址
     */
    private static String path = "C:/Users/Thinkpad/Documents/ITC TVP (ORMS) july 2020/ORMS-Copy.xlsx";
    // private static String path =
    // "/Users/Thinkpad/test_base/src/main/resources/static/ORMS-COPY.xlsx";

    /**
     * 匯出的Excel模板地址
     */
    private static String export = "C:/Users/Thinkpad/test_base/src/main/resources/static/ORMS-Copy-Export.xlsx";

    private static String pdfexport = "C:/Users/Thinkpad/test_base/ORMS-Copy-Export.pdf";

    private static String pdfexportd =
    "C:\\Users\\Thinkpad\\Downloads\\Excel_2.pdf";

    public static void main(String[] args) throws Exception {

        // [simple create / read & modify]
        // try (
        // // 讀入檔案流
        // FileInputStream inputStream = new FileInputStream(new File(path));
        // // 建立Excel物件
        // Workbook wb = new XSSFWorkbook(inputStream);) {
        // // 一般Excel都會有一個預設的sheet，所以直接獲取第一個sheet
        // Sheet sheet = wb.getSheetAt(0);
        // // 建立第一行
        // Row row = sheet.createRow(0);
        // // 設定A1的值
        // row.createCell(0).setCellValue("A1");
        // // 輸出檔案
        // FileOutputStream out = new FileOutputStream(export);
        // wb.write(out);
        // System.out.println("----------------輸出成功");
        // } catch (Exception e) {
        // // TODO: handle exception
        // System.out.println("error: "+ e);
        // }

        // [create & write with in code]
        // XSSFWorkbook workbook = new XSSFWorkbook();
        // XSSFSheet sheet = workbook.createSheet("Dummy min form");
        // // Object[][] datatypes = {
        // // {"Datatype", "Type", "Size(in bytes)"},
        // // {"int", "Primitive", 2},
        // // {"float", "Primitive", 4},
        // // {"double", "Primitive", 8},
        // // {"char", "Primitive", 1},
        // // {"String", "Non-Primitive", "No fixed size"}
        // // };
        // DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
        // LocalDateTime now = LocalDateTime.now();

        // CellStyle dateCellStyle = workbook.createCellStyle();
        // dateCellStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd"));

        // CellStyle popStyle = workbook.createCellStyle();
        // short format = (short)BuiltinFormats.getBuiltinFormat("#,##0.00");
        // popStyle.setDataFormat(format);
        // // popStyle.setFillBackgroundColor(bg);
        // Object[][] header ={
        // { "[SYSTEM]", "Staff:", "R78", "Title:","Sales" },
        // { "[SYSTEM]", "Ref:", 383838222 },{ "[SYSTEM]" },{},{}
        // } ;

        // Object[][] contents = { { "[SYSTEM]", "CLIENT:", "ABC", "Ada" }, {
        // "[SYSTEM]", "CONTACT:", "+852", "38888888" },
        // { "[SYSTEM]", "DATE:", dtf.format(now), "checking..." },{ "[SYSTEM]",
        // "Price:", 300 },
        // {"[SYSTEM]", "Unit: ", 4},{"[SYSTEM]", "Total Price: ", 1200, "0.0%"},
        // {"[SYSTEM]",
        // "Signature:"}, };

        // int rowNum = 0;
        // System.out.println("Creating excel");

        // for (Object[] h : header) {
        // Row row = sheet.createRow(rowNum++);
        // int colNum = 0;
        // for (Object field : h) {
        // Cell cell = row.createCell(colNum++);
        // if (field instanceof String) {

        // cell.setCellValue((String) field);
        // } else if (field instanceof Integer) {
        // // cell.setCellStyle(dateCellStyle);
        // cell.setCellValue((Integer) field);
        // }
        // }
        // }

        // for (Object[] content : contents) {
        // Row row = sheet.createRow(rowNum++);
        // int colNum = 0;
        // for (Object field : content) {
        // Cell cell = row.createCell(colNum++);
        // if (field instanceof String) {

        // cell.setCellValue((String) field);
        // } else if (field instanceof Integer) {
        // cell.setCellStyle(popStyle);
        // cell.setCellValue((Integer) field);
        // }
        // }
        // }

        // for (short i = sheet.getRow(0).getFirstCellNum(), end =
        // sheet.getRow(0).getLastCellNum(); i < end; i++) {
        // sheet.autoSizeColumn(i);
        // }

        // try {
        // FileOutputStream outputStream = new FileOutputStream(export);
        // workbook.write(outputStream);
        // workbook.close();
        // } catch (FileNotFoundException e) {
        // e.printStackTrace();
        // } catch (IOException e) {
        // e.printStackTrace();
        // }

        // System.out.println("Done");

        // [read from excel file & sys. print]
        // try (FileInputStream file = new FileInputStream(new File(path));

        // // Create Workbook instance holding reference to .xlsx file
        // XSSFWorkbook workbook = new XSSFWorkbook(file);)

        // {
        // DataFormatter formatter = new DataFormatter();
        // // Get first/desired sheet from the workbook
        // XSSFSheet sheet = workbook.getSheetAt(0);

        // // Iterate through each rows one by one
        // Iterator<Row> rowIterator = sheet.iterator();
        // int groupNo = 0;

        // // We will create output PDF document objects at this point
        // Document iText_xls_2_pdf = new Document();
        // PdfWriter.getInstance(iText_xls_2_pdf, new
        // FileOutputStream("Excel2PDF_Output.pdf"));
        // iText_xls_2_pdf.open();
        // // we have two columns in the Excel sheet, so we create a PDF table
        // // Note: There are ways to make this dynamic in nature, if you want to.
        // PdfPTable my_table = new PdfPTable(11);
        // // We will use the object below to dynamically add new data to the table
        // PdfPCell table_cell;
        // Phrase p;
        // Paragraph pa;
        // // table_cell.setBorder(Rectangle.NO_BORDER);
        // while (rowIterator.hasNext()) {
        // Row row = rowIterator.next();
        // // For each row, iterate through all the columns
        // Iterator<Cell> cellIterator = row.cellIterator();

        // int col = 0; // check the colume number

        // while (cellIterator.hasNext()) {
        // Cell cell = cellIterator.next();

        // // if (cell.getRowIndex() != 12 && col >= 12) {
        // // // skip the self-note columns except table header
        // // break;

        // if (col >= 12) {
        // // skip the self-note columns
        // break;
        // } else {

        // switch (cell.getCellType()) {

        // case NUMERIC:
        // System.out.print(formatter.formatCellValue(cell) + "\t");

        // table_cell = new PdfPCell(new Phrase(formatter.formatCellValue(cell)+" @n"));
        // // feel free to move the code below to suit to your needs
        // // table_cell.setIndent(5);
        // my_table.addCell(table_cell);

        // p = new Phrase(formatter.formatCellValue(cell) + " ");
        // iText_xls_2_pdf.add(p);

        // break;
        // case STRING:
        // if (cell.getStringCellValue().contains("[D-ITEM]")
        // || cell.getStringCellValue().contains("[SYSTEM]")
        // || cell.getStringCellValue().contains("[GROUP]")
        // || cell.getStringCellValue().contains("[SUBTOTAL]")
        // || cell.getStringCellValue().contains("[D-INFO]")
        // || cell.getStringCellValue().contains("[DISCOUNT]")
        // || cell.getStringCellValue().contains("[GRAND-TTL]")
        // || cell.getStringCellValue().contains("[D-INFO]")
        // || cell.getStringCellValue().contains("[TNC-HEAD]")
        // || cell.getStringCellValue().contains("[TNC]")
        // || cell.getStringCellValue().contains("[PAY-HEAD]")
        // || cell.getStringCellValue().contains("[PAY-HW]")
        // || cell.getStringCellValue().contains("[PAY-SVC]")) {
        // pa = new Paragraph("\n");
        // iText_xls_2_pdf.add(pa);
        // if (cell.getStringCellValue().contains("[GROUP]")) {
        // groupNo++;
        // System.out.println("Grouping");
        // }
        // break;
        // }

        // System.out.print(cell.getStringCellValue() + "\t");

        // // //Push the data from Excel to PDF Cell
        // table_cell = new PdfPCell(new Phrase(cell.getStringCellValue()+" @s"));
        // // //feel free to move the code below to suit to your needs
        // my_table.addCell(table_cell);

        // p = new Phrase(cell.getStringCellValue() + " ");
        // iText_xls_2_pdf.add(p);
        // break;
        // case BOOLEAN:
        // System.out.print(String.valueOf(cell.getBooleanCellValue()) + "\t");
        // break;
        // case FORMULA:
        // if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
        // // System.out.print(String.valueOf(cell.getNumericCellValue( ))+ "; ");
        // System.out.print(String.valueOf(
        // formatter.formatRawCellContents(cell.getNumericCellValue(), -1, "#,##0")
        // + "\t"));
        // // Push the data from Excel to PDF Cell
        // table_cell = new PdfPCell(new Phrase(
        // formatter.formatRawCellContents(cell.getNumericCellValue(), -1, "#,##0")+"
        // @f"));
        // // //feel free to move the code below to suit to your needs
        // my_table.addCell(table_cell);
        // p = new Phrase(String.valueOf(
        // formatter.formatRawCellContents(cell.getNumericCellValue(), -1, "#,##0")
        // + ""));
        // iText_xls_2_pdf.add(p);
        // break;
        // }
        // break;
        // default:
        // System.out.print("\t");
        // // Push the data from Excel to PDF Cell
        // table_cell = new PdfPCell(new Phrase("\t" +" @t"));
        // // //feel free to move the code below to suit to your needs
        // my_table.addCell(table_cell);
        // p = new Phrase(String.valueOf(" "));
        // iText_xls_2_pdf.add(p);
        // break;
        // }

        // col++; // count the column
        // }

        // // Check the cell type and format accordingly
        // }
        // System.out.println("\n");

        // file.close();
        // my_table.setWidthPercentage(100);

        // }
        // System.out.println("group(s): " + groupNo);

        // // Finally add the table to PDF document
        // iText_xls_2_pdf.add(my_table);
        // iText_xls_2_pdf.close();

        // } catch (Exception e) {
        // e.printStackTrace();
        // }

        // //[Excel to Pdf basic SPIRE]
        Locale.setDefault(new Locale("en-US")); // since lib is from .NET, Locale lang must be en-US
        Workbook workbook = new Workbook();
        workbook.loadFromFile((path));
        Worksheet sheet = workbook.getWorksheets().get(0);

        // Date -> generate date on specfic cell
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd");
        LocalDateTime now = LocalDateTime.now();
        sheet.setCellValue(5, 3, dtf.format(now));

        // Company address -> [note]cannot align right dual to CHAR(10) space
        sheet.getCellRange("I4").getCellStyle().setHorizontalAlignment(HorizontalAlignType.Left);
        sheet.getCellRange("I4").getCellStyle().setIndentLevel(10);
        //
        sheet.getCellRange("I4").getCellStyle().setVerticalAlignment(VerticalAlignType.Top);

        sheet.getCellRange("B1:" + "L" + sheet.getLastRow()).getCellStyle().setKnownColor(ExcelColors.White ); 
        
        // print range: column B to column L ->[note] no tag and hidden price table
        // ->row 1 to last row
        sheet.getPageSetup().setPrintArea("B1:" + "L" + sheet.getLastRow());
        // Fit to page
        workbook.getConverterSetting().setSheetFitToPage(true);


        // [NEED IMPROVE]copy item to database
        System.out.println("L: " + sheet.getRows().length);
        ArrayList<List<String>> dataList = new ArrayList<>();
        ArrayList<List<String>> TNCDataList = new ArrayList<>();
        ArrayList<List<String>> payDataList = new ArrayList<>();

        int item = 0;
        int TNCItem = 0;
        int payItem = 0;
        for (int row = 1; row <= sheet.getRows().length; row++) {

            System.out.println("R: " + row);
            System.out.println(sheet.getCellRange(row, 1).getText());

            if (sheet.getCellRange(row, 1).getText().equals("[D-ITEM]")) {
                List<String> dataRow = new ArrayList<>();
                for (int col = 2; col <= 12; col++) {
                    if (sheet.getCellRange(row, col).hasFormula()) {
                        dataRow.add(sheet.getCellRange(row, col).getFormulaValue().toString());
                        // System.out.println(sheet.getFormulaValue().toString());
                    } else if (!sheet.getCellRange(row, col).isBlank()) {
                        dataRow.add(sheet.getCellRange(row, col).getValue());
                        // System.out.println(sheet.getCellRange(row, col).getValue() );
                    }
                }
                dataList.add(dataRow);
                System.out.println("p: " + dataList.get(item).toString());
                item++;
            } else if (sheet.getCellRange(row, 1).getText().equals("[TNC]")) {
                List<String> TNCDataRow = new ArrayList<>();
                for (int col = 2; col <= 12; col++) {
                    if (!sheet.getCellRange(row, col).isBlank()) {
                        TNCDataRow.add(sheet.getCellRange(row, col).getValue());

                    }
                }
                TNCDataList.add(TNCDataRow);
                System.out.println("p: " + TNCDataList.get(TNCItem).toString());
                TNCItem++;
            } else if (sheet.getCellRange(row, 1).getText().matches("PAY")
                    && !sheet.getCellRange(row, 1).getText().matches("HEAD")) {
                List<String> payDataRow = new ArrayList<>();
                for (int col = 2; col <= 12; col++) {
                    if (!sheet.getCellRange(row, col).isBlank()) {
                        payDataRow.add(sheet.getCellRange(row, col).getValue());

                    }
                }
                payDataList.add(payDataRow);
                System.out.println("p: " + payDataList.get(payItem).toString());
                payItem++;
            }
        }
        // / end [NEED IMPROVE]copy item to database
        
        // Save as PDF document
        // workbook.saveToFile("ExcelToPDF.pdf", FileFormat.PDF);
        // workbook.saveToFile("Excel.pdf", FileFormat.PDF);
        workbook.saveToFile(pdfexportd, FileFormat.PDF);

        // FileOutputStream outputStream = new FileOutputStream(pdfexport);
        // workbook.saveToStream(outputStream,  FileFormat.PDF);
        // System.out.println("r: "+workbook.isSaved());
        

    }

}