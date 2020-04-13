/**
 * Copyright 2018 eShares, Inc. dba Carta, Inc.
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * https://www.apache.org/licenses/LICENSE-2.0
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
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


