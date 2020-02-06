package com.carta.yaml

import com.fasterxml.jackson.core.`type`.TypeReference

object CellType extends Enumeration {
  type CellType = Value

  val date, double, string, long = Value
}

class CellTypeReference extends TypeReference[CellType.type]

