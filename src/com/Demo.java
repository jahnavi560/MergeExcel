package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Vector;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo {

    private static Vector vectorDataExcelXLSX = new Vector(); 
    static XSSFSheet xssfSheet;
    static Sheet sh;
    static Vector<Vector> ParentVector = new Vector<Vector>();
    static FileOutputStream fos = null;
    public static File[] listFileNames;
    public static ArrayList <String> fileNames = new ArrayList<String>();

    public static String[] fileNames(String directoryPath) {
        File dir = new File(directoryPath);

        if (dir.isDirectory()) {
            listFileNames = dir.listFiles();
            for (File file : listFileNames) {
                if (file.isFile()) {
                    fileNames.add(file.getName());
                }
            }
        }
        return fileNames.toArray(new String[] {});
    }

    public static Vector readDataExcelXLSX(String fileName, int SheetNumber) {
        Vector vectorData = new Vector();
        String value="";
        try {
            FileInputStream fileInputStream = new FileInputStream(fileName);
            XSSFWorkbook xssfWorkBook = new XSSFWorkbook(fileInputStream);
            XSSFSheet xssfSheet = xssfWorkBook.getSheetAt(SheetNumber);
            Iterator rowIteration = xssfSheet.rowIterator();
            while (rowIteration.hasNext()) {
                XSSFRow xssfRow = (XSSFRow) rowIteration.next();
                Iterator cellIteration = xssfRow.cellIterator();
                Vector vectorCellEachRowData = new Vector();
                while (cellIteration.hasNext()) {
                    XSSFCell xssfCell = (XSSFCell) cellIteration.next();
              //      xssfCell.setCellType(Cell.CELL_TYPE_STRING);
                    value=xssfCell.toString();
                    if(value.contains(",")){
                        value=xssfCell.toString().replaceAll(",", "__");
                    }else{
                    }
                    System.out.println(value);
                    vectorCellEachRowData.addElement(value);
                }
                vectorData.addElement(vectorCellEachRowData);
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return vectorData;
    }

    public static void writeDataExcelXLSX(Vector<Vector> vectorData, Workbook wb, String sheetName) throws IOException {
        Row row;
        Cell cell;
        String[] cellvalue = null;
        int t = 0, p = 0;
        sh = wb.createSheet(sheetName);
        for (int i = 0; i < vectorData.size(); i++) {
            Vector vectorCellEachRowData = (Vector) vectorData.get(i);
            if (sheetName.equals("Time Sheet") )
                if (i != 0)    p = 1;
            for (int j = p; j < vectorCellEachRowData.size(); j++) {
                String str1 = vectorCellEachRowData.get(j).toString().replace("[", "");
                cellvalue = str1.split(",");
                row = sh.createRow((short) t);
                t++;
                for (int k = 0; k < cellvalue.length; k++) {
                    cell = row.createCell((short) k);
                    if (cellvalue[k].contains("__")) {
                        cellvalue[k] = cellvalue[k].replace("__", ",");
                    } else {
                    }
                    cell.setCellValue(cellvalue[k].replace("]", "").trim());
                }
            }
            row = sh.createRow((short) t);
            t++;
        }
    }

    public static void creatingMasterXlsxForAllModules() throws Exception {
        File file;
        Workbook wb = new XSSFWorkbook();
        fos = new FileOutputStream(  "E:\\Master.xlsx"); // Output File after merging all excel files
         file = new File("E:\\Sheets");
    
        String[] str = file.list();
        String[] tabName = { "Time Sheet"};
        for (int var = 0; var < 1; var++) {
            for (String st : str) {
                vectorDataExcelXLSX = readDataExcelXLSX("E:\\Sheets\\" + st, var);
                ParentVector.add(vectorDataExcelXLSX);
            }
            writeDataExcelXLSX(ParentVector, wb, tabName[var]);
            ParentVector.clear();
        }
        wb.write(fos);
        fos.close();
        System.out.println("Excel file has been generated!");
    }

    
    public static void main(String[] args) throws Exception {
            fileNames("E:\\Sheets");   // input folder where our excel files are present 
            creatingMasterXlsxForAllModules();
    }
}