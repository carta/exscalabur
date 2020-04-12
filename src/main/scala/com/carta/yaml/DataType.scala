package com.carta.yaml

import com.fasterxml.jackson.core.`type`.TypeReference

object DataType extends Enumeration {
  type DataCellType = Value
  val string, double, long = Value
}

class DataTypeReference extends TypeReference[DataType.type]