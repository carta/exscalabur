package com.carta.excel

import org.scalatest.Assertion
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

class CellFormulaParserSpec extends AnyFlatSpec with Matchers {
  "shiftRowNums" should "not change formula when shift factor is 0" in testShiftedFormula(
    testFormula = "SUM(A4, B4:B8)",
    expectedFormula = "SUM(A4, B4:B8)",
    shiftFactor = 0
  )

  "shiftRowNums" should "change formula when shift factor is positive" in testShiftedFormula(
    testFormula = "SUM(A4, B4:B8)",
    expectedFormula = "SUM(A8, B8:B12)",
    shiftFactor = 4
  )

  "shiftRowNums" should "change formula when shift factor is negative" in testShiftedFormula(
    testFormula = "SUM(A4, B4:B8)",
    expectedFormula = "SUM(A2, B2:B6)",
    shiftFactor = -2
  )

  "shiftRowNums" should "change formula with nested operations" in testShiftedFormula(
    testFormula = """SUM(A1, AVERAGE(B1, C2, D3), IF(OR(C2="red",C2="blue"),1,2))""",
    expectedFormula = """SUM(A5, AVERAGE(B5, C6, D7), IF(OR(C6="red",C6="blue"),1,2))""",
    shiftFactor = 4
  )

  def testShiftedFormula(testFormula: String, expectedFormula: String, shiftFactor: Int): Assertion = {
    val cellFormulaParser = new CellFormulaParser
    val shiftedFormula = cellFormulaParser.shiftRowNums(testFormula, shiftFactor)
    shiftedFormula shouldBe expectedFormula
  }
}
