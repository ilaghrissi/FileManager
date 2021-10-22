package com.tutorials.FileManager.services;
import java.io.FileInputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

public class ExcelManagerService {


  public void importExcelFile_XLXS_XSSF(String filepath) throws Exception {
    XSSFWorkbook wb = new XSSFWorkbook (new FileInputStream(filepath));

    try {
      System.out.println("Data dump:\n");

      for (int k = 0; k < wb.getNumberOfSheets(); k++) {
        XSSFSheet sheet = wb.getSheetAt(k);
        int rows = sheet.getPhysicalNumberOfRows();
        System.out.println("Sheet " + k + " \"" + wb.getSheetName(k) + "\" has " + rows
            + " row(s).");
        for (int r = 0; r < rows; r++) {
          XSSFRow row = sheet.getRow(r);
          if (row == null) {
            continue;
          }

          System.out.println("\nROW " + row.getRowNum() + " has " + row.getPhysicalNumberOfCells() + " cell(s).");
          for (int c = 0; c < row.getLastCellNum(); c++) {
            XSSFCell cell = row.getCell(c);
            String value;

            if (cell != null) {
              switch (cell.getCellTypeEnum()) {

                case FORMULA:
                  value = "FORMULA value=" + cell.getCellFormula();
                  break;

                case NUMERIC:
                  value = "NUMERIC value=" + cell.getNumericCellValue();
                  break;

                case STRING:
                  value = "STRING value=" + cell.getStringCellValue();
                  break;

                case BLANK:
                  value = "";
                  break;

                case BOOLEAN:
                  value = "BOOLEAN value-" + cell.getBooleanCellValue();
                  break;

                case ERROR:
                  value = "ERROR value=" + cell.getErrorCellValue();
                  break;

                default:
                  value = "UNKNOWN value of type " + cell.getCellTypeEnum();
              }
              System.out.println("CELL col=" + cell.getColumnIndex() + " VALUE="
                  + value);
            }
          }
        }
      }
    } finally {
      wb.close();
    }
  }





  public void importExcelFile_HSSF(String filepath) throws Exception {
    HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(filepath));

    try {
      System.out.println("Data dump:\n");

      for (int k = 0; k < wb.getNumberOfSheets(); k++) {
        HSSFSheet sheet = wb.getSheetAt(k);
        int rows = sheet.getPhysicalNumberOfRows();
        System.out.println("Sheet " + k + " \"" + wb.getSheetName(k) + "\" has " + rows
            + " row(s).");
        for (int r = 0; r < rows; r++) {
          HSSFRow row = sheet.getRow(r);
          if (row == null) {
            continue;
          }

          System.out.println("\nROW " + row.getRowNum() + " has " + row.getPhysicalNumberOfCells() + " cell(s).");
          for (int c = 0; c < row.getLastCellNum(); c++) {
            HSSFCell cell = row.getCell(c);
            String value;

            if (cell != null) {
              switch (cell.getCellTypeEnum()) {

                case FORMULA:
                  value = "FORMULA value=" + cell.getCellFormula();
                  break;

                case NUMERIC:
                  value = "NUMERIC value=" + cell.getNumericCellValue();
                  break;

                case STRING:
                  value = "STRING value=" + cell.getStringCellValue();
                  break;

                case BLANK:
                  value = "";
                  break;

                case BOOLEAN:
                  value = "BOOLEAN value-" + cell.getBooleanCellValue();
                  break;

                case ERROR:
                  value = "ERROR value=" + cell.getErrorCellValue();
                  break;

                default:
                  value = "UNKNOWN value of type " + cell.getCellTypeEnum();
              }
              System.out.println("CELL col=" + cell.getColumnIndex() + " VALUE="
                  + value);
            }
          }
        }
      }
    } finally {
      wb.close();
    }
  }
}
