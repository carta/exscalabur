package com.carta.excel

import java.util.Date

sealed trait CellValue

final case class CellString(s: String) extends CellValue

final case class CellDouble(d: Double) extends CellValue

final case class CellDate(d: Date) extends CellValue

final case class CellBoolean(b: Boolean) extends CellValue

final case class CellBlank() extends CellValue