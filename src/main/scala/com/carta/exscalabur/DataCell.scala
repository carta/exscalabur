package com.carta.exscalabur

case class DataCell(key: String, value: CellType) {
  def asTuple: (String, CellType) = key -> value
}

object DataCell {
  def apply(key: String, value: String): DataCell = apply(key, StringCellType(value))
  def apply(key: String, value: Long): DataCell = apply(key, LongCellType(value))
  def apply(key: String, value: Double): DataCell = apply(key, DoubleCellType(value))
}
