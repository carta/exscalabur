package test

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}

import main.{CellDouble, CellString, ExportModelUtils}
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.{XSSFCell, XSSFRow, XSSFSheet, XSSFWorkbook}
import org.scalatest.{FlatSpec, Matchers}
import test.ExcelTestHelpers._

import scala.collection.immutable
import scala.util.Random
import resource.managed

class ExcelWorkbookSpec extends FlatSpec with Matchers {
  "ExcelWorkbook" should "copy cellStyles from template workbook to output workbook" in {
    val templateName = "cellStyles"
    val templateStream: ByteArrayOutputStream = new ByteArrayOutputStream()
    val actualStream: ByteArrayOutputStream = new ByteArrayOutputStream()
    val templateWorkbook: XSSFWorkbook = new XSSFWorkbook()

    val sheet: XSSFSheet = templateWorkbook.createSheet(templateName)
    val row: XSSFRow = sheet.createRow(0)

    // Add multiple cell styles to the template
    val cellStyles = for (i <- 0 until 10) yield ExcelTestHelpers.randomCellStyle(templateWorkbook)
    for (colIndex <- 0 until 10) {
      val cell: XSSFCell = row.createCell(colIndex)
      cell.setCellStyle(cellStyles(colIndex))
      cell.setCellType(CellType.STRING)
      cell.setCellValue(Random.nextString(5))
    }
    templateWorkbook.write(templateStream)

    // Template stream also acts as expected stream since no substitutions are made
    writeOutputAndVerify(templateName, templateStream, templateStream, actualStream, None)

    ExcelTestHelpers.assertEqualsWorkbooks(
      new ByteArrayInputStream(templateStream.toByteArray()),
      new ByteArrayInputStream(actualStream.toByteArray())
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
    for (colIndex <- 0 until 10) {
      val cell: XSSFCell = row.createCell(colIndex)
      cell.setCellStyle(cellStyle)
      cell.setCellType(CellType.STRING)
      cell.setCellValue(Random.nextString(5))
    }
    templateWorkbook.write(templateStream)

    writeOutputAndVerify(templateName, templateStream, templateStream, actualStream, None)

    // Verify that a new cellStyle was not created for every cell
    // Should be 2 because one default cellStyle and one cellStyle added from the template workbook
    for (result <- managed(new XSSFWorkbook(new ByteArrayInputStream(actualStream.toByteArray())))) {
      result.getNumCellStyles should equal(2)
    }
  }

  "ExcelWorkbook" should "copy column widths from template workbook to output workbook" in {
    val templateName = "columnWidths"
    val templateStream: ByteArrayOutputStream = new ByteArrayOutputStream()
    val actualStream: ByteArrayOutputStream = new ByteArrayOutputStream()
    val templateWorkbook: XSSFWorkbook = new XSSFWorkbook()

    val sheet: XSSFSheet = templateWorkbook.createSheet(templateName)
    val row: XSSFRow = sheet.createRow(0)

    val cellStyle = ExcelTestHelpers.randomCellStyle(templateWorkbook)
    for (colIndex <- 0 until 10) {
      // Column width is set at the sheet level
      sheet.setColumnWidth(colIndex, (colIndex + 1) * 100)
    }
    templateWorkbook.write(templateStream)

    writeOutputAndVerify(templateName, templateStream, templateStream, actualStream, None)

    // Assert that each column width in the result sheet is as expected
    for (templateWorkbook <- managed(new XSSFWorkbook(new ByteArrayInputStream(actualStream.toByteArray())))) {
      for (colIndex <- 0 until 10) {
        sheet.getColumnWidth(colIndex) should equal ((colIndex + 1) * 100)
      }
    }
  }
}
