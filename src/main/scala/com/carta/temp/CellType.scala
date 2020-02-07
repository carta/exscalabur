package com.carta.temp


trait CellType

case class StringCellType(value: String) extends CellType

case class DoubleCellType(value: Double) extends CellType

case class LongCellType(value: Long) extends CellType