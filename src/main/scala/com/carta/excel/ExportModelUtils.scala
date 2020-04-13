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
package com.carta.excel

import java.time.Instant
import java.util.Date
import scala.util.control.Exception.allCatch

// TODO: Inline or rename or something
object ExportModelUtils {
  val SUBSTITUTION_KEY = "$KEY"
  val REPEATED_FIELD_KEY = "$REP"

  type ModelMap = Map[String, CellValue]

  def toCellStringFromString(string: String): CellValue = CellString(string)

  def toCellDouble(num: Number): CellValue = CellDouble(num.doubleValue)

  def toCellString(value: Any): CellValue = CellString(value.toString)

  def toCellDoubleFromString(num: String): CellValue = {
    (allCatch opt num.toDouble)
      .map(CellDouble)
      .getOrElse(CellString(num))
  }


  // to CellDouble Option converters
  def toCellDoubleFromDouble(double: Double): CellValue = CellDouble(double)

  def toCellDoubleFromLong(long: Long): CellValue = CellDouble(long.toDouble)

  // to CellDate Option converters
  def toCellDateFromLong(epochMillis: Number): CellDate =
  // epochMillis * hours * minutes * seconds * milliseconds
    CellDate(Date.from(Instant.ofEpochMilli(epochMillis.longValue * 24 * 60 * 60 * 1000)))

  def toCellDateFromTimestampMillis(epochMillis: Long, offsetSeconds: Long = 0): CellDate =
    CellDate(new Date(epochMillis + offsetSeconds * 1000))
}
