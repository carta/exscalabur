package com.carta.yaml

import com.fasterxml.jackson.module.scala.JsonScalaEnumeration

case class KeyObject
(
  keyType: KeyType.Value,
  columnType: CellType.Value,
  excelType: CellType.Value
)

class KeyObjectBuilder(
                        @JsonScalaEnumeration(classOf[KeyTypeReference])
                        var keyType: KeyType.Value,

                        @JsonScalaEnumeration(classOf[CellTypeReference])
                        columnType: CellType.Value,

                        @JsonScalaEnumeration(classOf[CellTypeReference])
                        excelType: CellType.Value) {
  def build(): KeyObject = {
    KeyObject(keyType, columnType, excelType)
  }
}

