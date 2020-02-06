package yaml

import com.fasterxml.jackson.core.`type`.TypeReference

object CellType extends Enumeration {
  type CellType = Value

  val date = Value
  val double = Value
  val string = Value
  val long = Value
}

class CellTypeReference extends TypeReference[CellType.type]

