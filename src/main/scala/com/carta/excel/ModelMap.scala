package com.carta.excel

import com.carta.excel.ExportModelUtils.ModelMap
import com.carta.exscalabur
import com.carta.exscalabur.{DataCell, DoubleCellType, LongCellType, StringCellType}
import com.carta.yaml.{YamlCellType, YamlEntry}

object ModelMap {
  def apply(keyType: String, dataRow: Iterable[DataCell], schema: Map[String, YamlEntry]): ModelMap = {

    dataRow.map(_.asTuple).map { case (key: String, value: exscalabur.CellType) =>
      val newKey = s"${keyType}.$key"
      val newValue = schema(newKey) match {
        // TODO different input output types
        case YamlEntry(_, YamlCellType.string, YamlCellType.string) =>
          ExportModelUtils.toCellStringFromString(value.asInstanceOf[StringCellType].value)
        case YamlEntry(_, YamlCellType.double, YamlCellType.double) =>
          ExportModelUtils.toCellDoubleFromDouble(value.asInstanceOf[DoubleCellType].value)
        case YamlEntry(_, YamlCellType.long, YamlCellType.double) =>
          ExportModelUtils.toCellDoubleFromLong(value.asInstanceOf[LongCellType].value)
        case YamlEntry(_, YamlCellType.long, YamlCellType.date) =>
          ExportModelUtils.toCellDateFromLong(value.asInstanceOf[LongCellType].value)
      }
      newKey -> newValue
    }.toMap
  }
}
