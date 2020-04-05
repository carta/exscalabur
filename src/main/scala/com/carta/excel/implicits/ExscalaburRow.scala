package com.carta.excel.implicits

import com.carta.excel.ExportModelUtils
import com.carta.excel.ExportModelUtils.ModelMap
import org.apache.poi.ss.usermodel.{CellType, Row}
import com.carta.excel.implicits.ExtendedRow._

import scala.collection.JavaConverters._

abstract class ExscalaburRow

class StaticRow extends ExscalaburRow

class RepeatedRow extends ExscalaburRow

class ImmutableRow extends ExscalaburRow

//object ExscalaburRow {
//
//  def apply(row: Row, repeatedData: Seq[ModelMap], staticData: ModelMap): Seq[ExscalaburRow] = {
//    if(row.isRepeatedRow) {
//      repeatedData.map(modelMap => row.)
//    }
//
//  }
//}
