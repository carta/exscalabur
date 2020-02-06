package com.carta.excel

import java.util.Date

abstract class CellValue

case class CellString(s: String) extends CellValue
case class CellDouble(d: Double) extends CellValue
case class CellDate(d: Date) extends CellValue
case class CellBoolean(b: Boolean) extends CellValue