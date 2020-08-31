package com.itcray.test_base;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Iterator;

import javax.print.DocFlavor.STRING;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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

    public static void main(String[] args) throws Exception {

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

        // XSSFWorkbook workbook = new XSSFWorkbook();
        // XSSFSheet sheet = workbook.createSheet("Datatypes in Java");
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
        // dateCellStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd
        // HH:mm:ss"));

        // Object[][] datatypes = {
        // {"[SYSTEM]", "CLIENT:", "ABC"},
        // {"[SYSTEM]", "CONTACT:", "3888888"},
        // {"[SYSTEM]", "DATE:", dtf.format(now) }
        // };

        // int rowNum = 0;
        // System.out.println("Creating excel");

        // for (Object[] datatype : datatypes) {
        // Row row = sheet.createRow(rowNum++);
        // int colNum = 0;
        // for (Object field : datatype) {
        // Cell cell = row.createCell(colNum++);
        // if (field instanceof String) {
        // cell.setCellValue((String) field);
        // } else if (field instanceof Integer) {
        // cell.setCellValue((Integer) field);
        // }
        // }
        // }

        // try {
        // FileOutputStream outputStream = new FileOutputStream(path);
        // workbook.write(outputStream);
        // workbook.close();
        // } catch (FileNotFoundException e) {
        // e.printStackTrace();
        // } catch (IOException e) {
        // e.printStackTrace();
        // }

        // System.out.println("Done");

        try (FileInputStream file = new FileInputStream(new File(path));

                // Create Workbook instance holding reference to .xlsx file
                XSSFWorkbook workbook = new XSSFWorkbook(file);)

        {
            DataFormatter formatter = new DataFormatter();
            // Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            // Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            int groupNo = 0;
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                // For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                int col = 0; //check the colume number
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    if (col >= 12) {
                        // skip the self-note columns
                        break;
                    } else {

                        switch (cell.getCellType()) {

                            case NUMERIC:
                                // System.out.print(formatter.formatCellValue(cell) + "\t");
                                System.out.print(formatter.formatCellValue(cell) + "; ");
                                break;
                            case STRING:
                                // System.out.print(cell.getStringCellValue() + "\t");
                                if (cell.getStringCellValue().contains("[GROUP]")) {
                                    groupNo++;
                                }
                                System.out.print(cell.getStringCellValue() + "; ");
                                break;
                            case BOOLEAN:
                                System.out.print(String.valueOf(cell.getBooleanCellValue()) + "; ");
                                break;
                            case FORMULA:
                                // System.out.print(String.valueOf(cell.getCellFormula()));
                                if (cell.getCachedFormulaResultType() == CellType.STRING) {
                                    System.out.print(cell.getRichStringCellValue() + "; ");
                                    break;
                                }
                                if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
                                    // System.out.print(String.valueOf(cell.getNumericCellValue( ))+ "; ");
                                    System.out.print(String.valueOf(
                                            formatter.formatRawCellContents(cell.getNumericCellValue(), -1, "#,##0")
                                                    + "; "));
                                    break;
                                }
                                break;
                            default:
                                System.out.print("\t");
                                break;
                        }
                        col++;
                    }
                    // Check the cell type and format accordingly
                }
                System.out.println("\n");

                file.close();
            }
            System.out.println("group(s): " + groupNo);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}