package com.carta.excel.implicits

import org.apache.poi.ss.usermodel.{Row, Sheet}

object ExtendedSheet {

  implicit class ExtendedSheet(sheet: Sheet) {
    def getRowIndices: Iterable[Int] = (0 to sheet.getLastRowNum).toVector

    def rowOpt(rowIndex: Int): Option[Row] = Option(sheet.getRow(rowIndex))
  }

}

