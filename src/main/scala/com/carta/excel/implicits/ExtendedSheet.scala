package com.carta.excel.implicits

import org.apache.poi.ss.usermodel.Sheet

object ExtendedSheet {

  implicit class ExtendedSheet(sheet: Sheet) {
    def getRowIndices: Seq[Int] = 0 to sheet.getLastRowNum
  }

}

