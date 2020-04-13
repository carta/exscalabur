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
package com.carta.excel

import com.carta.excel.ExportModelUtils.ModelMap
import com.carta.exscalabur
import com.carta.exscalabur.DataCell
import com.carta.yaml.{DataType, ExcelType, YamlEntry}

object ModelMap {
  def apply(keyType: String, dataRow: Iterable[DataCell], schema: Map[String, YamlEntry]): ModelMap = {

    dataRow.map(_.asTuple).map { case (key: String, cellValue: exscalabur.CellType) =>
      val newKey = s"${keyType}.$key"
      val newValue = schema(newKey) match {
        case YamlEntry(_, ExcelType.string) =>
          ExportModelUtils.toCellString(cellValue.getValueAsString)
        case YamlEntry(DataType.string, ExcelType.number) =>
          ExportModelUtils.toCellDoubleFromString(cellValue.getValueAsString)
        case YamlEntry(_, ExcelType.number) =>
          ExportModelUtils.toCellDouble(cellValue.getValueAsNumber)
        case YamlEntry(DataType.string, ExcelType.date) =>
          // TODO
          CellBlank()
        case YamlEntry(DataType.double, ExcelType.date) =>
          // TODO
          CellBlank()
        case YamlEntry(DataType.long, ExcelType.date) =>
          ExportModelUtils.toCellDateFromLong(cellValue.getValueAsNumber)
      }
      newKey -> newValue
    }.toMap
  }
}
