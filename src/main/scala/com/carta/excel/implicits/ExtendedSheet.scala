package com.carta.excel.implicits

import org.apache.poi.ss.usermodel.Sheet

object ExtendedSheet {

  implicit class ExtendedSheet(sheet: Sheet) {
    def getRowIndices: Range = 0 until sheet.getLastRowNum
  }

}

