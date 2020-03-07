package com.carta.excel

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}

import com.carta.excel.ExcelTestHelpers.{RepTemplateModel, SubTemplateModel, generateTestWorkbook, getModelSeq, writeOutputAndVerify}
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.{XSSFCell, XSSFRow, XSSFSheet, XSSFWorkbook}
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers
import resource.managed

import scala.util.Random


class ExcelWorkbookSpec extends AnyFlatSpec with Matchers {

  "ExcelWorkbook" should "copy cellStyles from template workbook to output workbook" in {
    val templateName = "cellStyles"
    val templateStream: ByteArrayOutputStream = new ByteArrayOutputStream()
    val actualStream: ByteArrayOutputStream = new ByteArrayOutputStream()
    val templateWorkbook: XSSFWorkbook = new XSSFWorkbook()

    val sheet: XSSFSheet = templateWorkbook.createSheet(templateName)
    val row: XSSFRow = sheet.createRow(0)

    (0 until 10).foreach(i => getRandomCell(templateWorkbook, row, i))

    templateWorkbook.write(templateStream)

    // Template stream also acts as expected stream since no substitutions are made
    writeOutputAndVerify(templateName, templateStream, templateStream, actualStream, None)

    ExcelTestHelpers.assertEqualsWorkbooks(
      new ByteArrayInputStream(templateStream.toByteArray),
      new ByteArrayInputStream(actualStream.toByteArray)
    )
  }

  "ExcelWorkbook" should "not create a new cellStyle for each cell in the output Workbook" in {
    val templateName = "singleCellStyle"
    val templateStream: ByteArrayOutputStream = new ByteArrayOutputStream()
    val actualStream: ByteArrayOutputStream = new ByteArrayOutputStream()
    val templateWorkbook: XSSFWorkbook = new XSSFWorkbook()

    val sheet: XSSFSheet = templateWorkbook.createSheet(templateName)
    val row: XSSFRow = sheet.createRow(0)

    // Only add 1 cell style to the template
    val cellStyle = ExcelTestHelpers.randomCellStyle(templateWorkbook)
    (0 until 10).foreach { i =>
      val cell = row.createCell(i)
      cell.setCellStyle(cellStyle)
      cell.setCellType(CellType.STRING)
      cell.setCellValue(Random.nextString(5))
    }

    templateWorkbook.write(templateStream)

    writeOutputAndVerify(templateName, templateStream, templateStream, actualStream, None)

    // Verify that a new cellStyle was not created for every cell
    // Should be 2 because one default cellStyle and one cellStyle added from the template workbook

    managed(toWorkbook(actualStream)).acquireAndGet { result =>
      result.getNumCellStyles shouldEqual 2
    }
  }

  "ExcelWorkbook" should "copy column widths from template workbook to output workbook" in {
    val templateName = "columnWidths"
    val templateStream: ByteArrayOutputStream = new ByteArrayOutputStream()
    val actualStream: ByteArrayOutputStream = new ByteArrayOutputStream()
    val templateWorkbook: XSSFWorkbook = new XSSFWorkbook()

    val sheet: XSSFSheet = templateWorkbook.createSheet(templateName)
    val row = sheet.createRow(0)

    // Column width is set at the sheet level
    (0 until 10).foreach { i =>
      row.createCell(i)
      sheet.setColumnWidth(i, (i + 1) * 100)
    }

    templateWorkbook.write(templateStream)

    writeOutputAndVerify(templateName, templateStream, templateStream, actualStream, None)

    // Assert that each column width in the result sheet is as expected
    managed(toWorkbook(actualStream))
      .acquireAndGet { workbook =>
        workbook.sheetIterator()
          .forEachRemaining { actualSheet =>
            (0 until 10).foreach { i =>
              actualSheet.getColumnWidth(i) should equal((i + 1) * 100)
            }
          }
      }
  }

  "ExcelWorkbook" should "do substitution from keys in template workbook to output workbook (happy path)" in {
    val templateName = "substitutionHappyPath"
    val templateStream: ByteArrayOutputStream = new ByteArrayOutputStream()
    val expectedStream: ByteArrayOutputStream = new ByteArrayOutputStream()
    val actualStream: ByteArrayOutputStream = new ByteArrayOutputStream()

    val model: SubTemplateModel = SubTemplateModel(Some("text"), Some(1234), Some(1234))
    val modelSeq = getModelSeq(model)

    // Build template
    val templateRows = Seq(
      // Row with static text, empty cell, and number
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("text")),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1234)),
      ),

      // Row with substitution keys
      modelSeq,
    )

    val templateWorkbook: XSSFWorkbook = generateTestWorkbook(templateName, templateRows, insertRowModelKeys = true)
    templateWorkbook.write(templateStream)
    val expectedWorkbook: XSSFWorkbook = generateTestWorkbook(templateName, templateRows)
    expectedWorkbook.write(expectedStream)

    writeOutputAndVerify(templateName, templateStream, expectedStream, actualStream, None, modelSeq.toMap)
  }

  "ExcelWorkbook" should "do substitution from keys in template workbook despite missing values in template" in {
    val templateName = "substitutionMissingValues"
    val templateStream: ByteArrayOutputStream = new ByteArrayOutputStream()
    val expectedStream: ByteArrayOutputStream = new ByteArrayOutputStream()
    val actualStream: ByteArrayOutputStream = new ByteArrayOutputStream()

    val model: SubTemplateModel = SubTemplateModel(Some("text"), Some(1234), Some(1234))
    val modelSeq = getModelSeq(model)

    // Build template
    val templateRows = Seq(
      // Row with static text, empty cell, and number
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("text")),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1234)),
      ),

      // Row with a substitution key missing
      modelSeq.drop(1)
    )

    val templateWorkbook: XSSFWorkbook = generateTestWorkbook(templateName, templateRows, insertRowModelKeys = true)
    templateWorkbook.write(templateStream)
    val expectedWorkbook: XSSFWorkbook = generateTestWorkbook(templateName, templateRows)
    expectedWorkbook.write(expectedStream)

    writeOutputAndVerify(templateName, templateStream, expectedStream, actualStream, None, modelSeq.toMap)
  }

  "ExcelWorkbook" should "do substitution from keys in template workbook despite Blank values in model" in {
    val templateName = "substitutionExtraKeys"
    val templateStream: ByteArrayOutputStream = new ByteArrayOutputStream()
    val expectedStream: ByteArrayOutputStream = new ByteArrayOutputStream()
    val actualStream: ByteArrayOutputStream = new ByteArrayOutputStream()

    val model: SubTemplateModel = SubTemplateModel(Some("text"), None, Some(1234))
    val modelSeq = getModelSeq(model)

    val templateRows = Seq(
      // Row with static text, empty cell, and number
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("text")),
        ("empty", CellBlank()),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1234)),
      ),

      // Row with substitution keys
      modelSeq
    )
    val expectedRows = Seq(
      // Row with static text, empty cell, and number
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("text")),
        ("empty", CellBlank()),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1234)),
      ),

      // Expect row with 1 None value missing
      modelSeq
    )

    val templateWorkbook: XSSFWorkbook = ExcelTestHelpers.generateTestWorkbook(templateName, templateRows, insertRowModelKeys = true)
    templateWorkbook.write(templateStream)
    val expectedWorkbook: XSSFWorkbook = ExcelTestHelpers.generateTestWorkbook(templateName, expectedRows)
    expectedWorkbook.write(expectedStream)

    writeOutputAndVerify(templateName, templateStream, expectedStream, actualStream, None, modelSeq.toMap)
  }

  "ExcelWorkbook" should "do substitution from keys in template workbook while ignoring extra keys in template" in {
    val templateName = "substitutionExtraKeys"
    val templateStream = new ByteArrayOutputStream()
    val expectedStream = new ByteArrayOutputStream()
    val actualStream = new ByteArrayOutputStream()

    val model = SubTemplateModel(Some("text"), Some(1234), Some(1234))
    val modelSeq = getModelSeq(model)

    // Build template
    val templateRows = Seq(
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("text")),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1234)),
      ),

      // Row with substitution keys
      modelSeq
    )
    val expectedRows = Seq(
      // Row with static text, empty cell, and number
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("text")),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1234)),
      ),


      // Expect row with 1 blank value
      (modelSeq.head._1, CellBlank()) +: modelSeq.drop(1).toList
    )

    ExcelTestHelpers.generateTestWorkbook(templateName, templateRows, insertRowModelKeys = true)
      .write(templateStream)

    ExcelTestHelpers.generateTestWorkbook(templateName, expectedRows)
      .write(expectedStream)

    writeOutputAndVerify(templateName, templateStream, expectedStream, actualStream, None, modelSeq.drop(1).toMap)
  }

  "ExcelWorkbook" should "load template for repeated rows and then insert set of repeated rows" in {
    val templateName = "substitutionExtraKeys"
    val templateStream: ByteArrayOutputStream = new ByteArrayOutputStream()
    val expectedStream: ByteArrayOutputStream = new ByteArrayOutputStream()
    val actualStream: ByteArrayOutputStream = new ByteArrayOutputStream()

    val repeatedModels = Seq(
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
    )
    val repeatedModelsSeq = repeatedModels.map(getModelSeq)

    val templateRows = Seq(
      // Row with static text, empty cell, and number
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("text")),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1234)),
      ),

      // 1 Row with repeated field keys
      repeatedModelsSeq.head
    )

    val expectedRows = Seq(
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("text")),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1234))
      ),
    ) ++ repeatedModelsSeq

    ExcelTestHelpers.generateTestWorkbook(templateName, templateRows, insertRowModelKeys = true)
      .write(templateStream)
    ExcelTestHelpers.generateTestWorkbook(templateName, expectedRows)
      .write(expectedStream)

    writeOutputAndVerify(templateName, templateStream, expectedStream, actualStream, Some(1), Map.empty, repeatedModelsSeq.map(_.toMap))
  }

  private def getRandomCell(templateWorkbook: XSSFWorkbook, row: XSSFRow, index: Int): XSSFCell = {
    val style = ExcelTestHelpers.randomCellStyle(templateWorkbook)
    val cell = row.createCell(index)
    cell.setCellStyle(style)
    cell.setCellType(CellType.STRING)
    cell.setCellValue(Random.nextString(5))
    cell
  }

  private def toWorkbook(stream: ByteArrayOutputStream): XSSFWorkbook = {
    new XSSFWorkbook(new ByteArrayInputStream(stream.toByteArray))
  }
}
