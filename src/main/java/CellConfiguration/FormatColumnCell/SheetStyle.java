package CellConfiguration.FormatColumnCell;

import lombok.Data;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
@Data
public class SheetStyle {

    private XSSFWorkbook workbook;
    private Cell cell;
    private Row row;
    private CellStyle headerStyle;
    private CellStyle textStyle;
    private CellStyle numberStyle;
    private XSSFSheet sheet;

    public SheetStyle(XSSFWorkbook workbook, XSSFSheet sheet, Cell cell, Row row){
        this.workbook = workbook;
        this.sheet = sheet;
        this.cell = cell;
        this.row = row;
        setLayoutDimensions();
        setHeader();
        setFont();
        setTextStyle();
        setNumberStyle();
    }
    public void setFont(){
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 15);
        font.setFontName("Courier New");
        font.setItalic(true);
        font.setStrikeout(true);
        font.setBold(true);
        headerStyle.setFont(font);
    }
    private void setHeader(){
        headerStyle =  this.workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.LEFT);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 6));
        sheet.createFreezePane(3,1);
    }
    private void setTextStyle(){
        textStyle =  this.workbook.createCellStyle();
        textStyle.setFillForegroundColor(IndexedColors.SEA_GREEN.getIndex());
        textStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        textStyle.setAlignment(HorizontalAlignment.CENTER);
        textStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        textStyle.setBorderBottom(BorderStyle.HAIR);
        textStyle.setBottomBorderColor(IndexedColors.BLACK1.getIndex());
        textStyle.setBorderLeft(BorderStyle.HAIR);
        textStyle.setLeftBorderColor(IndexedColors.GREEN.getIndex());
        textStyle.setBorderRight(BorderStyle.HAIR);
        textStyle.setRightBorderColor(IndexedColors.BLUE.getIndex());
        textStyle.setBorderTop(BorderStyle.HAIR);
        textStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
    }
    private void setNumberStyle(){
        numberStyle =  this.workbook.createCellStyle();
        XSSFDataFormat numberFormat = workbook.createDataFormat();
        numberStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        numberStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        numberStyle.setDataFormat(numberFormat.getFormat("#,##0.00"));
        numberStyle.setAlignment(HorizontalAlignment.CENTER);
        numberStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        sheet.createFreezePane(2,1);
    }
    private void setLayoutDimensions(){
        sheet.setDefaultColumnWidth(40);
        sheet.setDefaultRowHeight((short) 500);
    }


}
