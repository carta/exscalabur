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
