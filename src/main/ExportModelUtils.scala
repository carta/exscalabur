package main

import java.time.Instant
import java.util.Date

object ExportModelUtils {
  val SUBSTITUTION_KEY = "$KEY"
  val REPEATED_FIELD_KEY = "$REP"

  type ModelMap = Map[String, CellValue]

  def getModelMap(number: Long): ModelMap = {
    Map(
      (s"${SUBSTITUTION_KEY}.number" -> toCellDoubleFromLong(number))
      )
  }

  def getModelMap(string: String): ModelMap = {
    Map(
      (s"${REPEATED_FIELD_KEY}.string" -> toCellStringFromString(string))
    )
  }

  // to CellString Option converters
  def toCellStringFromString(string: String) = CellString(string)

  // to CellDouble Option converters
  def toCellDoubleFromDouble(double: Double) = CellDouble(double)

  def toCellDoubleFromLong(long: Long) = CellDouble(long.toDouble)

  // to CellDate Option converters
  def toCellDateFromLong(epochMillis: Long) =
    // epochMillis * hours * minutes * seconds * milliseconds
    CellDate(Date.from(Instant.ofEpochMilli(epochMillis * 24 * 60 * 60 * 1000)))

  def toCellDateFromTimestampMillis(epochMillis: Long, offsetSeconds: Long = 0): CellDate =
    CellDate(new Date(epochMillis + offsetSeconds * 1000))
}
