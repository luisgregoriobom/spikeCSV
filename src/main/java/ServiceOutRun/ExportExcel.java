package ServiceOutRun;

import CellConfiguration.FormatColumnCell.SheetStyle;
import CellConfiguration.ImportDataCell.DataManipulation;
import Model.Product;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExportExcel {

    public static List<Product> getProductListMock(){
        List<Product> productsList = new ArrayList<>();
        productsList.add(new Product(1l,  "Produto 1",  200.5d));
        productsList.add(new Product(2l,  "Produto 2",  1050.5d));
        productsList.add(new Product(3l,  "Produto 3",  50d));
        productsList.add(new Product(4l,  "Produto 4",  200d));
        productsList.add(new Product(5l,  "Produto 5",  450d));
        productsList.add(new Product(6l,  "Produto 6",  150.5d));
        productsList.add(new Product(7l,  "Produto 7",  300.99d));
        productsList.add(new Product(8l,  "Produto 8",  1000d));
        productsList.add(new Product(9l,  "Produto 9",  350d));
        productsList.add(new Product(10l, "Produto 10", 200d));
        productsList.add(new Product(11l, "Produto 11", 350d));
        productsList.add(new Product(12l, "Produto 12", 850d));
        productsList.add(new Product(13l, "Produto 13", 320d));
        productsList.add(new Product(14l, "Produto 14", 400d));
        productsList.add(new Product(16l, "Produto 16", 150d));
        productsList.add(new Product(15l, "Produto 15", 200d));
        productsList.add(new Product(17l, "Produto 17", 80d));
        productsList.add(new Product(19l, "Produto 19", 120d));
        productsList.add(new Product(20l, "Produto 20", 10d));
        productsList.add(new Product(21l, "Produto 21", 800d));
        productsList.add(new Product(22l, "Produto 22", 1000d));
        productsList.add(new Product(23l, "Produto 23", 1500d));
        return productsList;
    }

    public static void main(String[] args) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Model.Products");
        Cell cell = null;
        Row row = null;

        try {
            SheetStyle sheetStyle = new SheetStyle(workbook, sheet, cell, row);
            DataManipulation dataManipulation = new DataManipulation(workbook, sheet, cell, row, sheetStyle);
            dataManipulation.insertData(getProductListMock(), "PAC ABRIL - Validade: 01/04/2022 a 30/04/2022");
            exportSheet("/home/develcode02/products.xlsx", workbook);

            System.out.println("FUNCIONOU CARAIO!!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    private static void  exportSheet(String path, XSSFWorkbook workbook) throws IOException {
        FileOutputStream out = new FileOutputStream(new File(path));
        workbook.write(out);
        out.close();
        workbook.close();
    }
}