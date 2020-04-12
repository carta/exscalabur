package com.carta.exscalabur

sealed trait CellType {
  def getValueAsString: String
  def getValueAsNumber: Number
}

final case class StringCellType(value: String) extends CellType {
  override def getValueAsString: String = value

  override def getValueAsNumber: Number = throw new UnsupportedOperationException
}

final case class DoubleCellType(value: Double) extends CellType {
  override def getValueAsString: String = value.toString

  override def getValueAsNumber: Number = value
}

final case class LongCellType(value: Long) extends CellType {
  override def getValueAsString: String = value.toString
  override def getValueAsNumber: Number = value
}