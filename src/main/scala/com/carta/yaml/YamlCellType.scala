package com.carta.yaml

import com.fasterxml.jackson.core.`type`.TypeReference

object YamlCellType extends Enumeration {
  type YamlCellType = Value

  val date, double, string, long = Value
}

class CellTypeReference extends TypeReference[YamlCellType.type]

