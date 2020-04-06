package com.carta.excel.implicits

import com.carta.excel.ExportModelUtils
import org.apache.poi.ss.usermodel.{Cell, CellType, Row}

import scala.collection.JavaConverters._

object ExtendedRow {

  implicit class ExtendedRow(row: Row) {
    def isRepeatedRow: Boolean = rowHasCellStartingWithKey(ExportModelUtils.REPEATED_FIELD_KEY)

    def isStaticRow: Boolean = rowHasCellStartingWithKey(ExportModelUtils.SUBSTITUTION_KEY)

    private def rowHasCellStartingWithKey(key: String): Boolean = {
      row.cellIterator().asScala
        .filter(cell => cell.getCellType == CellType.STRING)
        .map(cell => cell.getStringCellValue)
        .exists(cellValue => cellValue.startsWith(key))
    }
  }

}
