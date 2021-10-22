package com.tutorials.FileManager.services;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

class ExcelManagerServiceTest {

  @Test
  void import_XLXS_File() throws Exception {

    ExcelManagerService excelManagerService = new ExcelManagerService();
    excelManagerService.importExcelFile_XLXS_XSSF("D://text.xlsx");
  }
}