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
