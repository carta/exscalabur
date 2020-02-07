//package com.carta.excel
//
//import java.io.{ByteArrayInputStream, ByteArrayOutputStream, File, InputStream}
//import java.util.Date
//
//import org.apache.poi.ss.usermodel.{BorderStyle, CellStyle, CellType}
//import org.apache.poi.xssf.usermodel._
//import org.scalatest.{Assertion, Matchers, Succeeded}
//import resource.managed
//
//import scala.collection.immutable
//import scala.util.{Failure, Random, Success}
//
//object ExcelTestHelpers extends Matchers {
//  /*
//   * Utility methods
//   */
//  def createTempFile(prefix: String): File = {
//    val f: File = File.createTempFile(prefix, ".xlsx")
//    f.deleteOnExit()
//    return f
//  }
//
//  /*
//   * Assertion methods
//   */
//  def assertEqualsWorkbooks(expectedWorkbookStream: InputStream, actualWorkbookStream: InputStream): Assertion = {
//    managed(new XSSFWorkbook(expectedWorkbookStream)) and managed(new XSSFWorkbook(actualWorkbookStream)) acquireAndGet (entry => entry match {
//      case (expectedWorkbook, actualWorkbook) =>
//        assert(expectedWorkbook.getNumberOfSheets() == actualWorkbook.getNumberOfSheets())
//        assert(List.range(0, expectedWorkbook.getNumberOfSheets()) map { sheetIndex =>
//          assertEqualsSheets(
//            Option(expectedWorkbook.getSheetAt(sheetIndex)),
//            Option(actualWorkbook.getSheetAt(sheetIndex)),
//          )
//        } forall (_ == Succeeded))
//    })
//  }
//
//  def assertEqualsSheets(expectedSheetOpt: Option[XSSFSheet], actualSheetOpt: Option[XSSFSheet]): Assertion = {
//    expectedSheetOpt match {
//      case None => assert(actualSheetOpt.isEmpty)
//
//      case Some(expectedSheet) =>
//        assert(!actualSheetOpt.isEmpty)
//
//        val actualSheet = actualSheetOpt.get
//
//        assert(expectedSheet.getSheetName() == actualSheet.getSheetName())
//        assert(expectedSheet.getLastRowNum() == actualSheet.getLastRowNum())
//
//        // XSSFSheet::getLastRowNum returns the index of the last row not the number of rows
//        assert(List.range(0, expectedSheet.getLastRowNum() + 1) map { rowIndex =>
//          assertEqualsRows(
//            Option(expectedSheet.getRow(rowIndex)),
//            Option(actualSheet.getRow(rowIndex)),
//          )
//        } forall (_ == Succeeded))
//    }
//  }
//
//  def assertEqualsRows(expectedRowOpt: Option[XSSFRow], actualRowOpt: Option[XSSFRow]): Assertion = {
//    expectedRowOpt match {
//      case None => assert(actualRowOpt.isEmpty)
//
//      case Some(expectedRow) =>
//        assert(!actualRowOpt.isEmpty, s"Row ${expectedRow.getRowNum} in sheet ${expectedRow.getSheet.getSheetName}  should not be empty")
//
//        val actualRow = actualRowOpt.get
//
//        assert(expectedRow.getLastCellNum() == actualRow.getLastCellNum(), s"Row ${actualRow.getRowNum} in sheet ${actualRow.getSheet.getSheetName} should have the same number of columns")
//
//        assert(List.range(0, expectedRow.getLastCellNum()) map { cellIndex =>
//          val expectedCell = Option(expectedRow.getCell(cellIndex))
//          val actualCell = Option(actualRow.getCell(cellIndex))
//
//          (assertEqualsCellStyles(expectedCell, actualCell), assertEqualsCellValues(expectedCell, actualCell))
//        } forall (assertionTuple => assertionTuple._1 == Succeeded && assertionTuple._2 == Succeeded))
//    }
//  }
//
//  def assertEqualsCellStyles(expectedCellOpt: Option[XSSFCell], actualCellOpt: Option[XSSFCell]): Assertion = {
//    expectedCellOpt match {
//      case None =>
//        actualCellOpt match {
//          case None => assert(actualCellOpt.isEmpty)
//          case Some(actualCell) => assert(actualCellOpt.isEmpty, errorLocation(actualCell))
//        }
//      case Some(expectedCell: XSSFCell) =>
//        assert(!actualCellOpt.isEmpty, errorLocation(expectedCell))
//
//        val actualCell = actualCellOpt.get
//
//        val expectedStyle = expectedCell.getCellStyle
//        val actualStyle = actualCell.getCellStyle
//
//        assertEqualsFonts(expectedStyle.getFont, actualStyle.getFont, actualCell)
//        assertEqualsBorders(expectedStyle, actualStyle, actualCell)
//        expectedStyle.getFillForegroundColor should equal(actualStyle.getFillForegroundColor)
//        expectedStyle.getAlignment should equal(actualStyle.getAlignment)
//        expectedStyle.getVerticalAlignment should equal(actualStyle.getVerticalAlignment)
//    }
//  }
//
//  def assertEqualsCellValues(expectedCellOpt: Option[XSSFCell], actualCellOpt: Option[XSSFCell]): Assertion = {
//    expectedCellOpt match {
//      case None => assert(actualCellOpt.isEmpty)
//
//      case Some(expectedCell: XSSFCell) =>
//        assert(!actualCellOpt.isEmpty, errorLocation(expectedCell))
//
//        val actualCell = actualCellOpt.get
//
//        assert(expectedCell.getCellType() == actualCell.getCellType(), errorLocation(actualCell))
//        expectedCell.getCellType match {
//          case CellType.NUMERIC =>
//            assert(expectedCell.getNumericCellValue() == actualCell.getNumericCellValue(), errorLocation(actualCell))
//          case CellType.STRING =>
//            assert(expectedCell.getStringCellValue() == actualCell.getStringCellValue(), errorLocation(actualCell))
//          case CellType.BOOLEAN =>
//            assert(expectedCell.getBooleanCellValue() == actualCell.getBooleanCellValue(), errorLocation(actualCell))
//          case CellType.BLANK => assert(true)
//          case _ => fail(s"Attempting to assert equality of unsupported cell value type")
//        }
//    }
//  }
//
//  private def assertEqualsFonts(expectedFont: XSSFFont, actualFont: XSSFFont, actualCell:XSSFCell): Assertion = {
//    assert(expectedFont.getFamily == actualFont.getFamily, errorLocation(actualCell))
//    assert(expectedFont.getColor == actualFont.getColor, errorLocation(actualCell))
//    assert(expectedFont.getBold == actualFont.getBold, errorLocation(actualCell))
//    assert(expectedFont.getItalic == actualFont.getItalic, errorLocation(actualCell))
//    assert(expectedFont.getUnderline == actualFont.getUnderline, errorLocation(actualCell))
//    assert(expectedFont.getFontHeight == actualFont.getFontHeight, errorLocation(actualCell))
//  }
//
//  private def assertEqualsBorders(expectedStyle: XSSFCellStyle, actualStyle: XSSFCellStyle, actualCell:XSSFCell): Assertion = {
//    assert(expectedStyle.getBorderBottom == actualStyle.getBorderBottom, errorLocation(actualCell))
//    assert(expectedStyle.getBorderLeft == actualStyle.getBorderLeft, errorLocation(actualCell))
//    assert(expectedStyle.getBorderRight == actualStyle.getBorderRight, errorLocation(actualCell))
//    assert(expectedStyle.getBorderTop == actualStyle.getBorderTop, errorLocation(actualCell))
//  }
//
//  private def errorLocation(actualCell: XSSFCell) = {
//    s"Error occurred at row ${actualCell.getRowIndex} and column ${actualCell.getColumnIndex} in sheet ${actualCell.getSheet.getSheetName}"
//  }
//
//  /*
//   * Excel template generation helper methods
//   */
//  def addRow(workbook: XSSFWorkbook, rowIndex: Int, cellValues: Seq[Option[CellValue]]): XSSFRow = {
//    val sheet: XSSFSheet = workbook.getSheetAt(0)
//    val row: XSSFRow = sheet.createRow(rowIndex)
//
//    for (colIndex <- 0 until cellValues.length) {
//      val cell: XSSFCell = row.createCell(colIndex)
//
//      cell.setCellStyle(randomCellStyle(workbook))
//
//      cellValues(colIndex) match {
//        case Some(v: CellValue) => v match {
//          case CellString(s) =>
//            cell.setCellType(CellType.STRING)
//            cell.setCellValue(s)
//          case CellDouble(d) =>
//            cell.setCellType(CellType.NUMERIC)
//            cell.setCellValue(d)
//          case CellDate(d) =>
//            cell.setCellType(CellType.STRING)
//            cell.setCellValue(d)
//          case CellBoolean(b) =>
//            cell.setCellType(CellType.BOOLEAN)
//            cell.setCellValue(b)
//        }
//        case None => Unit
//    }
//  }
//
//  row
//}
//
//def randomCellStyle(workbook: XSSFWorkbook): CellStyle = {
//  val style = workbook.createCellStyle()
//
//    // Set a bunch of random styles that will be copied over
//    style.setBorderTop(BorderStyle.values()(randomInt(BorderStyle.values.length)))
//    style.setBorderRight(BorderStyle.values()(randomInt(BorderStyle.values.length)))
//    style.setBorderBottom(BorderStyle.values()(randomInt(BorderStyle.values.length)))
//    style.setBorderLeft(BorderStyle.values()(randomInt(BorderStyle.values.length)))
//    style.setFillBackgroundColor(new XSSFColor(randomRGB(), new DefaultIndexedColorMap()))
//
//    style
//  }
//
//  def randomInt(maxExclusive: Int): Int = Random.nextInt(maxExclusive)
//  def randomRGB(): Array[Byte] = Array[Byte](
//    randomInt(0xF).asInstanceOf[Byte],
//    randomInt(0xF).asInstanceOf[Byte],
//    randomInt(0xF).asInstanceOf[Byte]
//  )
//
//  // Useful keys for adding certain cells that aren't supported by the substitution or repeated field keys
//  val STATIC_TEXT_KEY = "$TEST_STATIC_TEXT"
//  val STATIC_NUMBER_KEY = "$TEST_STATIC_NUMBER"
//
//  // Ordered version of ExportModelUtils.ModelMap to make it easier to test for exact cell locations
//  type ModelSeq = immutable.Seq[(String, CellValue)]
//
//  // generateTestWorkbook creates a new workbook with 1 sheet represented by the given rowModels.
//  //
//  // Depending on the value of insertRowModelKeys, either ExportModelUtils.SUBSTITUTION_KEY or ExportModelUtils.REPEATED_FIELD_KEY
//  // will be inserted when encountered or the values associated with those keys instead.
//  // This is useful for generating a template workbook which would contain the keys or an expected value workbook
//  // which would contain the corresponding values instead that are expected.
//  def generateTestWorkbook(name: String, rowModels: immutable.Seq[ModelSeq], insertRowModelKeys: Boolean = false): XSSFWorkbook = {
//    val workbook: XSSFWorkbook = new XSSFWorkbook()
//    val sheet: XSSFSheet = workbook.createSheet(name)
//
//    for {
//      rowIndex <- 0 until rowModels.length
//      templateRow = sheet.createRow(rowIndex)
//      rowModel = rowModels(rowIndex)
//    } populateRow(rowModel, templateRow, insertRowModelKeys)
//
//    workbook
//  }
//
//  private def populateRow(rowModel: ModelSeq, row: XSSFRow, substituteKey: Boolean): Unit = {
//    rowModel.foldLeft(0) {
//      (cellIndex, keyValue) =>
//        val cell = row.createCell(cellIndex)
//
//        keyValue._1 match {
//          case key if key.startsWith(STATIC_TEXT_KEY) => keyValue._2 match {
//            case CellString(string: String) => cell.setCellValue(string)
//            case _ => cell.setCellType(CellType.BLANK)
//          }
//
//          case key if key.startsWith(STATIC_NUMBER_KEY) => keyValue._2 match {
//            case CellDouble(double: Double) => cell.setCellValue(double)
//            case _ => cell.setCellType(CellType.BLANK)
//          }
//
//          case key if key.startsWith(ExportModelUtils.SUBSTITUTION_KEY) || key.startsWith(ExportModelUtils.REPEATED_FIELD_KEY) =>
//            if (substituteKey) cell.setCellValue(key)
//            else keyValue._2 match {
//              case CellString(string: String) => cell.setCellValue(string)
//              case CellDouble(double: Double) => cell.setCellValue(double)
//              case CellBoolean(boolean: Boolean) => cell.setCellValue(boolean)
//              case CellDate(date: Date) => cell.setCellValue(date)
//              case _ => cell.setCellType(CellType.BLANK)
//            }
//
//          case _ => cell.setCellType(CellType.BLANK)
//        }
//
//        cellIndex + 1
//    }
//  }
//
//  // writeOutputAndVerify creates and writes the contents of the given ByteArrayOutputStreams into an ExcelWorkbook
//  // and then verifies that output of this ExcelWorkbook is as expected
//  def writeOutputAndVerify(
//    templateName: String,
//    templateStream: ByteArrayOutputStream,
//    expectedStream: ByteArrayOutputStream,
//    actualStream: ByteArrayOutputStream,
//    expectedCopyResult: Option[Int],
//    modelSeq: ExcelTestHelpers.ModelSeq = immutable.Seq(),
//    repeatedModelsSeq: immutable.Seq[ExcelTestHelpers.ModelSeq] = immutable.Seq(),
//    batchSize: Int = 1000,
//  ): Unit = {
//    val templateStreamMap = Map((templateName, managed(new ByteArrayInputStream(templateStream.toByteArray()))))
//
//    for {
//      workbook <- managed(new ExcelWorkbook(templateStreamMap, 10))
//    } {
//      val rowIndexOpt = workbook.copyAndSubstitute(templateName, modelSeq.toMap)
//
//      rowIndexOpt should equal(expectedCopyResult)
//      rowIndexOpt map { startRowIndex =>
//        var rowIndex = startRowIndex
//
//        for (batch <- repeatedModelsSeq.grouped(batchSize)) {
//          workbook.insertRows(templateName, startRowIndex, rowIndex, batch map { modelSeq => modelSeq.toMap }) match {
//            case Success(newRowIndex) =>
//              newRowIndex should equal(rowIndex + batch.length)
//              rowIndex = newRowIndex
//            case Failure(exception) =>
//              fail("Expected to receive row index of next set of rows to insert but got None")
//
//          }
//        }
//      }
//      workbook.write(actualStream)
//    }
//
//    val expectedInputStream = new ByteArrayInputStream(expectedStream.toByteArray())
//    val actualInputStream = new ByteArrayInputStream(actualStream.toByteArray())
//
//    ExcelTestHelpers.assertEqualsWorkbooks(expectedInputStream, actualInputStream)
//  }
//
//  /*
//   * Test models to test substitution and repeated templates
//   */
//  case class SubTemplateModel(
//    stringField: String,
//    longField: Long,
//    doubleField: Double
//  )
//
//  case class RepTemplateModel(
//    stringField: String,
//    longField: Long,
//    doubleField: Double
//  )
//
//  def getModelSeq(t: SubTemplateModel): ExcelTestHelpers.ModelSeq = {
//    immutable.Seq(
//      (s"${ExportModelUtils.SUBSTITUTION_KEY}_string_field", ExportModelUtils.toCellStringFromString(t.stringField)),
//      (s"${ExportModelUtils.SUBSTITUTION_KEY}_long_field", ExportModelUtils.toCellDoubleFromLong(t.longField)),
//      (s"${ExportModelUtils.SUBSTITUTION_KEY}_double_field", ExportModelUtils.toCellDoubleFromDouble(t.doubleField)),
//    )
//  }
//
//  def getModelSeq(r: RepTemplateModel): ExcelTestHelpers.ModelSeq = {
//    immutable.Seq(
//      (s"${ExportModelUtils.REPEATED_FIELD_KEY}_string_field", ExportModelUtils.toCellStringFromString(r.stringField)),
//      (s"${ExportModelUtils.REPEATED_FIELD_KEY}_long_field", ExportModelUtils.toCellDoubleFromLong(r.longField)),
//      (s"${ExportModelUtils.REPEATED_FIELD_KEY}_double_field", ExportModelUtils.toCellDoubleFromDouble(r.doubleField)),
//    )
//  }
//}
