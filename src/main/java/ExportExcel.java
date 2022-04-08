import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExportExcel {

    public static void main(String[] args) {

        //variáveis
        int rownum = 0;
        int cellnum = 0;
        Cell cell;
        Row row;

        // Criando o arquivo e uma planilha chamada "Product"
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Product");

        //Configurando código para criar estilos de células (Cores, alinhamento, formatação, etc..)
        XSSFDataFormat numerFormat = workbook.createDataFormat();

        // Definindo alguns padroes de layout
        sheet.setDefaultColumnWidth(40);
        sheet.setDefaultRowHeight((short) 500);

        //Carregando Produtos para Lista
        List<Product> productsList = new ArrayList<>();
        productsList.add(new Product(1l, "Produto 1", 200.5d));
        productsList.add(new Product(2l, "Produto 2", 1050.5d));
        productsList.add(new Product(3l, "Produto 3", 50d));
        productsList.add(new Product(4l, "Produto 4", 200d));
        productsList.add(new Product(5l, "Produto 5", 450d));
        productsList.add(new Product(6l, "Produto 6", 150.5d));
        productsList.add(new Product(7l, "Produto 7", 300.99d));
        productsList.add(new Product(8l, "Produto 8", 1000d));
        productsList.add(new Product(9l, "Produto 9", 350d));
        productsList.add(new Product(10l, "Produto 10", 200d));


        //Estilização das Colunas
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        sheet.createFreezePane(0,1);


        CellStyle textStyle = workbook.createCellStyle();
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
        sheet.createFreezePane(2,1);

        CellStyle numberStyle = workbook.createCellStyle();
        numberStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        numberStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        numberStyle.setDataFormat(numerFormat.getFormat("#,##0.00"));
        numberStyle.setAlignment(HorizontalAlignment.CENTER);
        numberStyle.setVerticalAlignment(VerticalAlignment.CENTER);


        //Configurando Header de Colunas, iniciar contagem de célula sempre no 0.
        row = sheet.createRow(rownum++);
        cell = row.createCell(cellnum++);


        //Tentando configurar scroll lock com formatação na fonte
        cell.setCellValue("PAC ABRIL - Validade: 01/04/2022 a 30/04/2022");
        cell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));



        //Adicionando os dados dos produtos na planilha
        for (Product product : productsList) {
            row = sheet.createRow(rownum++);
            cellnum = 0;

            cell = row.createCell(cellnum++);
            cell.setCellStyle(textStyle);
            cell.setCellValue(product.getId());

            cell = row.createCell(cellnum++);
            cell.setCellStyle(textStyle);
            cell.setCellValue(product.getName());

            cell = row.createCell(cellnum++);
            cell.setCellStyle(numberStyle);
            cell.setCellValue(product.getPrice());
        }

        try {

            //Escrevendo o arquivo em Disco.
            FileOutputStream out = new FileOutputStream(new File("/home/develcode02/products.xlsx"));
            workbook.write(out);
            out.close();
            workbook.close();
            System.out.println("FUNCIONOU CARAIO!!");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
        }
    }
}
