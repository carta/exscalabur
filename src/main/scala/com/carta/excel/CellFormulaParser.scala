package com.carta.excel

import scala.util.matching.Regex

class CellFormulaParser {
  val cellReferenceRegex: Regex = "([A-Z])([0-9])".r

  def shiftRowNums(formula: String, shiftFactor: Int): String = {
    val cellRefs = cellReferenceRegex.findAllIn(formula)
      .toList
      .mapConserve {
        case cellReferenceRegex(row, col) => f"$row${col.toInt + shiftFactor}"
        case token: String => token
      }

    formula.split(cellReferenceRegex.toString)
      .zipAll(cellRefs, "", "")
      .map { case (a, b) => a + b }
      .mkString
  }
}
