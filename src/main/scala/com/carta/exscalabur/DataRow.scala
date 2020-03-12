package com.carta.exscalabur

import scala.collection.immutable.VectorBuilder
import scala.collection.mutable

class DataRow(val data: Map[String, CellType]) {
  def getCells: Iterable[DataCell] = data.map { case (key, value) => DataCell(key, value) }
}

object DataRow {

  class Builder() {
    private val data = new VectorBuilder[DataCell]
    private val dataMap = mutable.Map.empty[String, DataCell]

    def addCell(cell: DataCell): Builder = {
      data += cell
      this
    }

    def addAllCells(cells: Iterable[DataCell]): Builder = {
      cells.foreach(addCell)
      this
    }

    def build(): DataRow = new DataRow(data.result().map(cell => cell.asTuple).toMap)

    def addCell(key: String, value: CellType): Builder = addCell(DataCell(key, value))

    def addCell(key: String, value: Long): Builder = addCell(DataCell(key, value))

    def addCell(key: String, value: String): Builder = addCell(DataCell(key, value))

    def addCell(key: String, value: Double): Builder = addCell(DataCell(key, value))

  }

  object Builder {
    def apply() = new Builder
  }

}


