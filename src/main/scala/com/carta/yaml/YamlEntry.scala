package com.carta.yaml

import com.fasterxml.jackson.module.scala.JsonScalaEnumeration

case class YamlEntry
(
  keyType: KeyType.Value,
  columnType: YamlCellType.Value,
  excelType: YamlCellType.Value
)

class EntryBuilder(
                    @JsonScalaEnumeration(classOf[KeyTypeReference])
                        var keyType: KeyType.Value,

                    @JsonScalaEnumeration(classOf[CellTypeReference])
                        columnType: YamlCellType.Value,

                    @JsonScalaEnumeration(classOf[CellTypeReference])
                        excelType: YamlCellType.Value) {
  def build(): YamlEntry = {
    YamlEntry(keyType, columnType, excelType)
  }
}

