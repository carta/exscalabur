package com.carta.exscalabur

import scala.collection.mutable

class DataRow(val data: Map[String, CellType])

object DataRow {

  class Builder() {
    private val dataMap = mutable.Map.empty[String, CellType]

    def addCell(key: String, value: CellType): Builder = {
      dataMap.put(key, value)
      this
    }

    def addCell(key: String, value: Long): Builder = addCell(key, LongCellType(value))

    def addCell(key: String, value: String): Builder = addCell(key, StringCellType(value))

    def addCell(key: String, value: Double): Builder = addCell(key, DoubleCellType(value))

    def build(): DataRow = new DataRow(dataMap.toMap)
  }

  object Builder {
    def apply() = new Builder
  }

}


