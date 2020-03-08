package com.carta.excel

import org.apache.poi.xssf.usermodel.{XSSFSheet, XSSFWorkbook}

object WorkbookHelpers {

  implicit class ExtendedWorkbook(workbook: XSSFWorkbook) {
    def sheets(): IndexedSeq[XSSFSheet] = {
      (0 until workbook.getNumberOfSheets).map(workbook.getSheetAt)
    }
  }

}
