/**
 * Copyright 2018 eShares, Inc. dba Carta, Inc.
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * https://www.apache.org/licenses/LICENSE-2.0
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.carta.exscalabur

import java.time.{LocalDate, ZoneId}

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

final case class DateCellType(date: LocalDate, zoneId: ZoneId) extends CellType {
  override def getValueAsString: String = date.atStartOfDay(zoneId).toString

  override def getValueAsNumber: Number = date.atStartOfDay(zoneId).toInstant.toEpochMilli
}