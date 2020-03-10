package com.carta.yaml

import com.fasterxml.jackson.module.scala.JsonScalaEnumeration

case class YamlEntry
(
  keyType: KeyType.Value,
  columnType: CellType.Value,
  excelType: CellType.Value
)

class EntryBuilder(
                        @JsonScalaEnumeration(classOf[KeyTypeReference])
                        var keyType: KeyType.Value,

                        @JsonScalaEnumeration(classOf[CellTypeReference])
                        columnType: CellType.Value,

                        @JsonScalaEnumeration(classOf[CellTypeReference])
                        excelType: CellType.Value) {
  def build(): YamlEntry = {
    YamlEntry(keyType, columnType, excelType)
  }
}
