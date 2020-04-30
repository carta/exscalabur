package com.carta.compat.java.exscalabur

import java.time.{LocalDate, ZoneId}

import com.carta.exscalabur._

class DataCell(key: String, value: CellType) {
  def this(key: String, value: String) = {
    this(key, StringCellType(value))
  }

  def this(key: String, value: Long) = {
    this(key, LongCellType(value))
  }

  def this(key: String, value: Double) = {
    this(key, DoubleCellType(value))
  }

  def this(key: String, date: LocalDate, zoneId: ZoneId = ZoneId.systemDefault()) = {
    this(key, DateCellType(date, zoneId))
  }

  def asScala: com.carta.exscalabur.DataCell = {
    com.carta.exscalabur.DataCell(key, value)
  }

}
