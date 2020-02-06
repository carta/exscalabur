package com.carta.yaml

import com.fasterxml.jackson.core.`type`.TypeReference

object KeyType extends Enumeration {
  type KeyType = Value
  val repeated, single = Value
}

class KeyTypeReference extends TypeReference[KeyType.type]
