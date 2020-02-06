package com.carta.yaml

import com.fasterxml.jackson.module.scala.JsonScalaEnumeration

case class KeyObject
(
  key: String,
  keyType: KeyType.Value,
  columnType: CellType.Value,
  excelType: CellType.Value
)

class KeyObjectBuilder(var key: String,

                       @JsonScalaEnumeration(classOf[KeyTypeReference])
                       var keyType: KeyType.Value,

                       @JsonScalaEnumeration(classOf[CellTypeReference])
                       columnType: CellType.Value,

                       @JsonScalaEnumeration(classOf[CellTypeReference])
                       excelType: CellType.Value) {
  def build(): KeyObject = {
    KeyObject(key, keyType, columnType, excelType)
  }
}

