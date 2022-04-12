package CellConfiguration.ImportDataCell;

import CellConfiguration.FormatColumnCell.SheetStyle;
import Model.Product;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;

public class DataManipulation {

    private XSSFWorkbook workbook;
    private Cell cell;
    private Row row;
    private XSSFSheet sheet;

    private SheetStyle sheetStyle;

    public DataManipulation (XSSFWorkbook workbook, XSSFSheet sheet, Cell cell, Row row, SheetStyle sheetStyle) {
        this.workbook = workbook;
        this.sheet = sheet;
        this.cell = cell;
        this.row = row;
        this.sheetStyle = sheetStyle;
    }

    public void insertData(List<Product> productList, String title){
        int rownum = 0;
        int cellnum = 0;
        row = sheet.createRow(rownum++);
        cell = row.createCell(cellnum++);
        cell.setCellStyle(sheetStyle.getHeaderStyle());
        cell.setCellValue(title);

        for (Product product : productList) {
            row = sheet.createRow(rownum++);
            cellnum = 0;

            addCell(cellnum++, product.getId().toString(), sheetStyle.getTextStyle());
            addCell(cellnum++, product.getName(), sheetStyle.getTextStyle());
            addCell(cellnum++, product.getPrice().toString(), sheetStyle.getNumberStyle());
        }
    }
    private void addCell(Integer position, String value, CellStyle style){
        cell = row.createCell(position);
        cell.setCellStyle(style);
        cell.setCellValue(value);
    }

}
