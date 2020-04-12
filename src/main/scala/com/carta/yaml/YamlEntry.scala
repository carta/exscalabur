package com.carta.yaml

import com.fasterxml.jackson.module.scala.JsonScalaEnumeration

case class YamlEntry
(
  dataType: DataType.Value,
  excelType: ExcelType.Value
)

class EntryBuilder(
                    @JsonScalaEnumeration(classOf[DataTypeReference])
                    dataType: DataType.Value,

                    @JsonScalaEnumeration(classOf[ExcelTypeReference])
                    excelType: ExcelType.Value) {
  def build(): YamlEntry = {
    YamlEntry(dataType, excelType)
  }
}

