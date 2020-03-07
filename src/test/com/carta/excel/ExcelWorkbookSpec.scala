package com.carta.excel

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}

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
    ExcelTestHelpers.writeOutputAndVerify(templateName, templateStream, templateStream, actualStream, None)

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

    ExcelTestHelpers.writeOutputAndVerify(templateName, templateStream, templateStream, actualStream, None)

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

    ExcelTestHelpers.writeOutputAndVerify(templateName, templateStream, templateStream, actualStream, None)

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
