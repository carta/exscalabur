package com.carta.yaml

import com.fasterxml.jackson.core.`type`.TypeReference

object ExcelType extends Enumeration {
  type ExcelType = Value
  val string, number, date = Value
}

class ExcelTypeReference extends TypeReference[ExcelType.type]