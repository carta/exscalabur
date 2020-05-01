package com.carta.compat.java.exscalabur

import com.carta.exscalabur.CellType

class DataRow() {
  private val scalaDataRow = com.carta.exscalabur.DataRow()

  def addCell(cell: DataCell): DataRow = {
    scalaDataRow.addCell(cell.asScala)
    this
  }

  def addCell(key: String, value: CellType): DataRow = addCell(new DataCell(key, value))

  def addCell(key: String, value: Long): DataRow = addCell(new DataCell(key, value))

  def addCell(key: String, value: Double): DataRow = addCell(new DataCell(key, value))

  def addCell(key: String, value: String): DataRow = addCell(new DataCell(key, value))

  def asScala = scalaDataRow
}
