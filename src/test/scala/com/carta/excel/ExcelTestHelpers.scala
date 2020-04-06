package com.carta.excel

import java.io.{ByteArrayInputStream, ByteArrayOutputStream, File, InputStream}
import java.util.Date

import com.carta.excel.ExportModelUtils.ModelMap
import com.carta.yaml.{KeyType, YamlCellType, YamlEntry}
import org.apache.poi.ss.usermodel.{BorderStyle, CellStyle, CellType}
import org.apache.poi.xssf.usermodel._
import org.scalatest.matchers.should.Matchers
import org.scalatest.{Assertion, Succeeded}
import resource.managed

import scala.collection.JavaConverters._
import scala.util.{Failure, Random, Success}

object ExcelTestHelpers extends Matchers {
  // Useful keys for adding certain cells that aren't supported by the substitution or repeated field keys
  val STATIC_TEXT_KEY = "$TEST_STATIC_TEXT"
  val STATIC_NUMBER_KEY = "$TEST_STATIC_NUMBER"
  val STATIC_BOOL_KEY = "$TEST_STATIC_BOOL"
  val STATIC_DOUBLE_KEY = "$TEST_STATIC_DOUBLE"
  val STATIC_DATE_KEY = "$TEST_STATIC_DATE"

  /*
 * Excel template generation helper methods
 */
  def createTempFile(prefix: String): File = {
    val f = File.createTempFile(prefix, ".xlsx")
    f.deleteOnExit()
    f
  }

  def addRow(workbook: XSSFWorkbook, rowIndex: Int, cellValues: IndexedSeq[Option[CellValue]]): XSSFRow = {
    val sheet: XSSFSheet = workbook.getSheetAt(0)
    val row: XSSFRow = sheet.createRow(rowIndex)

    cellValues.indices.foreach { i =>
      val cell = row.createCell(i)
      cell.setCellStyle(randomCellStyle(workbook))

      cellValues(i).get match {
        case CellString(s) =>
          cell.setCellType(CellType.STRING)
          cell.setCellValue(s)
        case CellDouble(d) =>
          cell.setCellType(CellType.NUMERIC)
          cell.setCellValue(d)
        case CellDate(d) =>
          cell.setCellType(CellType.STRING)
          cell.setCellValue(d)
        case CellBoolean(b) =>
          cell.setCellType(CellType.BOOLEAN)
          cell.setCellValue(b)
      }
    }
    row
  }

  def randomCellStyle(workbook: XSSFWorkbook): CellStyle = {
    val style = workbook.createCellStyle()

    // Set a bunch of random styles that will be copied over
    style.setBorderTop(BorderStyle.values()(randomInt(BorderStyle.values.length)))
    style.setBorderRight(BorderStyle.values()(randomInt(BorderStyle.values.length)))
    style.setBorderBottom(BorderStyle.values()(randomInt(BorderStyle.values.length)))
    style.setBorderLeft(BorderStyle.values()(randomInt(BorderStyle.values.length)))
    style.setFillBackgroundColor(new XSSFColor(randomRGB(), new DefaultIndexedColorMap()))

    style
  }

  def randomInt(maxExclusive: Int): Int = Random.nextInt(maxExclusive)

  // generateTestWorkbook creates a new workbook with 1 sheet represented by the given rowModels.
  //
  // Depending on the value of insertRowModelKeys, either ExportModelUtils.SUBSTITUTION_KEY or ExportModelUtils.REPEATED_FIELD_KEY
  // will be inserted when encountered or the values associated with those keys instead.
  // This is useful for generating a template workbook which would contain the keys or an expected value workbook
  // which would contain the corresponding values instead that are expected.
  def generateTestWorkbook(name: String, rowModels: Seq[Iterable[(String, CellValue)]], insertRowModelKeys: Boolean = false): XSSFWorkbook = {
    val workbook = new XSSFWorkbook()
    val sheet = workbook.createSheet(name)

    rowModels.indices.foreach { i =>
      val templateRow = sheet.createRow(i)
      val rowModel = rowModels(i)
      populateRow(rowModel, templateRow, insertRowModelKeys)
    }

    workbook
  }

  def randomRGB(): Array[Byte] = Array[Byte](
    randomInt(0xF).asInstanceOf[Byte],
    randomInt(0xF).asInstanceOf[Byte],
    randomInt(0xF).asInstanceOf[Byte]
  )

  def getModelSchema(t: StaticTemplateModel): Map[String, YamlEntry] = {
    val schemaBuilder = Map.newBuilder[String, YamlEntry]
    if (t.stringField.isDefined) {
      val yamlEntry = YamlEntry(KeyType.single, YamlCellType.string, YamlCellType.string)
      schemaBuilder += s"${ExportModelUtils.SUBSTITUTION_KEY}.string_field" -> yamlEntry
    }
    if (t.doubleField.isDefined) {
      val yamlEntry = YamlEntry(KeyType.single, YamlCellType.double, YamlCellType.double)
      schemaBuilder += s"${ExportModelUtils.SUBSTITUTION_KEY}.double_field" -> yamlEntry
    }
    if (t.longField.isDefined) {
      val yamlEntry = YamlEntry(KeyType.single, YamlCellType.long, YamlCellType.double)
      schemaBuilder += s"${ExportModelUtils.SUBSTITUTION_KEY}.long_field" -> yamlEntry
    }
    schemaBuilder.result()
  }

  def getModelSeq(t: StaticTemplateModel): Iterable[(String, CellValue)] = {
    val modelSeqBuilder = List.newBuilder[(String, CellValue)]
    if (t.stringField.isDefined) {
      val cellString = ExportModelUtils.toCellStringFromString(t.stringField.get)
      modelSeqBuilder += ((s"${ExportModelUtils.SUBSTITUTION_KEY}.string_field", cellString))
    }
    if (t.doubleField.isDefined) {
      val cellDouble = ExportModelUtils.toCellDouble(t.doubleField.get)
      modelSeqBuilder += ((s"${ExportModelUtils.SUBSTITUTION_KEY}.double_field", cellDouble))
    }
    if (t.longField.isDefined) {
      val cellLong = ExportModelUtils.toCellDouble(t.longField.get)
      modelSeqBuilder += ((s"${ExportModelUtils.SUBSTITUTION_KEY}.long_field", cellLong))
    }

    modelSeqBuilder.result
  }

  def getModelSeq(r: RepTemplateModel): Iterable[(String, CellValue)] = {
    val modelSeqBuilder = List.newBuilder[(String, CellValue)]
    if (r.stringField.isDefined) {
      val cellString = ExportModelUtils.toCellStringFromString(r.stringField.get)
      modelSeqBuilder += ((s"${ExportModelUtils.REPEATED_FIELD_KEY}.string_field", cellString))
    }
    if (r.doubleField.isDefined) {
      val cellDouble = ExportModelUtils.toCellDoubleFromDouble(r.doubleField.get)
      modelSeqBuilder += ((s"${ExportModelUtils.REPEATED_FIELD_KEY}.long_field", cellDouble))
    }
    if (r.longField.isDefined) {
      val cellLong = ExportModelUtils.toCellDoubleFromLong(r.longField.get)
      modelSeqBuilder += ((s"${ExportModelUtils.REPEATED_FIELD_KEY}.double_field", cellLong))
    }

    modelSeqBuilder.result
  }

  /*
   * Assertion methods
   */
  def assertEqualsWorkbooks(expectedWorkbookStream: InputStream, actualWorkbookStream: InputStream): Assertion = {
    (managed(new XSSFWorkbook(expectedWorkbookStream)) and managed(new XSSFWorkbook(actualWorkbookStream)))
      .acquireAndGet {
        case (expectedWorkbook, actualWorkbook) =>
          //          assert(expectedWorkbook.getAllPictures.size() == actualWorkbook.getAllPictures.size())
          assert(expectedWorkbook.getNumberOfSheets == actualWorkbook.getNumberOfSheets)
          assert(List.range(0, expectedWorkbook.getNumberOfSheets) map { sheetIndex =>
            assertEqualsSheets(
              Option(expectedWorkbook.getSheetAt(sheetIndex)),
              Option(actualWorkbook.getSheetAt(sheetIndex)),
            )
          } forall (_ == Succeeded))
      }
  }

  def assertEqualsSheets(expectedSheetOpt: Option[XSSFSheet], actualSheetOpt: Option[XSSFSheet]): Assertion = {
    expectedSheetOpt match {
      case None => assert(actualSheetOpt.isEmpty)

      case Some(expectedSheet) =>
        assert(actualSheetOpt.isDefined)

        val actualSheet = actualSheetOpt.get

        assert(expectedSheet.getSheetName == actualSheet.getSheetName)
        assert(actualSheet.getLastRowNum == expectedSheet.getLastRowNum)

        // XSSFSheet::getLastRowNum returns the index of the last row not the number of rows
        assert(List.range(0, expectedSheet.getLastRowNum + 1) map { rowIndex =>
          assertEqualsRows(
            Option(expectedSheet.getRow(rowIndex)),
            Option(actualSheet.getRow(rowIndex)),
          )
        } forall (_ == Succeeded))
    }
  }

  def assertEqualsRows(expectedRowOpt: Option[XSSFRow], actualRowOpt: Option[XSSFRow]): Assertion = {
    expectedRowOpt match {
      case None => assert(actualRowOpt.isEmpty)

      case Some(expectedRow) =>
        assert(actualRowOpt.isDefined, s"Row ${expectedRow.getRowNum} in sheet ${expectedRow.getSheet.getSheetName}  should not be empty")

        val actualRow = actualRowOpt.get

        assert(expectedRow.getLastCellNum == actualRow.getLastCellNum, s"Row ${actualRow.getRowNum} in sheet ${actualRow.getSheet.getSheetName} should have the same number of columns")

        assert(List.range(0, expectedRow.getLastCellNum) map { cellIndex =>
          val expectedCell = Option(expectedRow.getCell(cellIndex))
          val actualCell = Option(actualRow.getCell(cellIndex))

          (assertEqualsCellStyles(expectedCell, actualCell), assertEqualsCellValues(expectedCell, actualCell))
        } forall (assertionTuple => assertionTuple._1 == Succeeded && assertionTuple._2 == Succeeded))
    }
  }

  def assertEqualsCellStyles(expectedCellOpt: Option[XSSFCell], actualCellOpt: Option[XSSFCell]): Assertion = {
    expectedCellOpt match {
      case None =>
        actualCellOpt match {
          case None => assert(actualCellOpt.isEmpty)
          case Some(actualCell) => assert(actualCellOpt.isEmpty, errorLocation(actualCell))
        }
      case Some(expectedCell: XSSFCell) =>
        assert(actualCellOpt.isDefined, errorLocation(expectedCell))

        val actualCell = actualCellOpt.get

        val expectedStyle = expectedCell.getCellStyle
        val actualStyle = actualCell.getCellStyle

        assertEqualsFonts(expectedStyle.getFont, actualStyle.getFont, actualCell)
        assertEqualsBorders(expectedStyle, actualStyle, actualCell)
        expectedStyle.getFillForegroundColor should equal(actualStyle.getFillForegroundColor)
        expectedStyle.getAlignment should equal(actualStyle.getAlignment)
        expectedStyle.getVerticalAlignment should equal(actualStyle.getVerticalAlignment)
    }
  }

  def assertEqualsCellValues(expectedCellOpt: Option[XSSFCell], actualCellOpt: Option[XSSFCell]): Assertion = {
    expectedCellOpt match {
      case None => assert(actualCellOpt.isEmpty)

      case Some(expectedCell: XSSFCell) =>
        assert(actualCellOpt.isDefined, errorLocation(expectedCell))

        val actualCell = actualCellOpt.get

        assert(expectedCell.getCellType == actualCell.getCellType, errorLocation(actualCell))
        expectedCell.getCellType match {
          case CellType.NUMERIC =>
            assert(expectedCell.getNumericCellValue == actualCell.getNumericCellValue, errorLocation(actualCell))
          case CellType.STRING =>
            assert(expectedCell.getStringCellValue == actualCell.getStringCellValue, errorLocation(actualCell))
          case CellType.BOOLEAN =>
            assert(expectedCell.getBooleanCellValue == actualCell.getBooleanCellValue, errorLocation(actualCell))
          case CellType.BLANK => assert(true)
          case _ => fail(s"Attempting to assert equality of unsupported cell value type")
        }
    }
  }

  def putPicture(sheet: XSSFSheet, picData: PictureData): Unit = {
    val workbook = sheet.getWorkbook
    val anchor = workbook.getCreationHelper.createClientAnchor()
    val drawing = sheet.createDrawingPatriarch()

    val imgStream = getClass.getResourceAsStream(picData.imgSrc)

    val picIndex = workbook.addPicture(imgStream, picData.imgType)
    anchor.setCol1(picData.posn.col1)
    anchor.setCol2(picData.posn.col2)
    anchor.setRow1(picData.posn.row1)
    anchor.setRow2(picData.posn.row2)
    anchor.setDx1(picData.posn.dx1)
    anchor.setDx2(picData.posn.dx2)
    anchor.setDy1(picData.posn.dy1)
    anchor.setDy2(picData.posn.dy2)

    drawing.createPicture(anchor, picIndex)
  }

  case class PictureData(imgSrc: String, imgType: Int, posn: PicturePosition)

  case class PicturePosition(col1: Int = 0,
                             col2: Int = 0,
                             row1: Int = 0,
                             row2: Int = 0,
                             dx1: Int = 0,
                             dx2: Int = 0,
                             dy1: Int = 0,
                             dy2: Int = 0)

  // writeOutputAndVerify creates and writes the contents of the given ByteArrayOutputStreams into an ExcelWorkbook
  // and then verifies that output of this ExcelWorkbook is as expected
  def writeOutputAndVerify(templateName: String,
                           templateStream: ByteArrayOutputStream,
                           expectedStream: ByteArrayOutputStream,
                           actualStream: ByteArrayOutputStream,
                           expectedCopyResult: Option[Int],
                           staticRowData: ModelMap = Map.empty,
                           repeatedRowData: Seq[ModelMap] = Seq.empty,
                           batchSize: Int = 1000,
                          ): Unit = {
    val templateBytes = managed(new ByteArrayInputStream(templateStream.toByteArray))
    val templateStreamMap = Map(templateName -> templateBytes)
    managed(new ExcelWorkbook(templateStreamMap, 10))
      .acquireAndGet { workbook =>
        workbook.sheets(templateName).foreach { templateSheet =>
          templateSheet.rowIterator().asScala
            .foldLeft(0) { (outputRow, row) =>
              workbook.insertRows(
                templateName,
                row.getRowNum,
                templateSheet.getSheetName,
                outputRow,
                staticRowData,
                repeatedRowData
              ) match {
                case Failure(e) => fail(e)
                case Success(nextRow) => nextRow
              }
            }
        }
        workbook.write(actualStream)
      }


    val expectedInputStream = new ByteArrayInputStream(expectedStream.toByteArray)
    val actualInputStream = new ByteArrayInputStream(actualStream.toByteArray)

    ExcelTestHelpers.assertEqualsWorkbooks(expectedInputStream, actualInputStream)
  }

  private def assertEqualsFonts(expectedFont: XSSFFont, actualFont: XSSFFont, actualCell: XSSFCell): Assertion = {
    assert(expectedFont.getFamily == actualFont.getFamily, errorLocation(actualCell))
    assert(expectedFont.getColor == actualFont.getColor, errorLocation(actualCell))
    assert(expectedFont.getBold == actualFont.getBold, errorLocation(actualCell))
    assert(expectedFont.getItalic == actualFont.getItalic, errorLocation(actualCell))
    assert(expectedFont.getUnderline == actualFont.getUnderline, errorLocation(actualCell))
    assert(expectedFont.getFontHeight == actualFont.getFontHeight, errorLocation(actualCell))
  }

  private def assertEqualsBorders(expectedStyle: XSSFCellStyle, actualStyle: XSSFCellStyle, actualCell: XSSFCell): Assertion = {
    assert(expectedStyle.getBorderBottom == actualStyle.getBorderBottom, errorLocation(actualCell))
    assert(expectedStyle.getBorderLeft == actualStyle.getBorderLeft, errorLocation(actualCell))
    assert(expectedStyle.getBorderRight == actualStyle.getBorderRight, errorLocation(actualCell))
    assert(expectedStyle.getBorderTop == actualStyle.getBorderTop, errorLocation(actualCell))
  }

  private def errorLocation(actualCell: XSSFCell) = {
    s"Error occurred at row ${actualCell.getRowIndex} and column ${actualCell.getColumnIndex} in sheet ${actualCell.getSheet.getSheetName}"
  }

  private def populateRow(rowModel: Iterable[(String, CellValue)], row: XSSFRow, substituteKey: Boolean): Unit = {
    rowModel.foldLeft(0) {
      (cellIndex, keyValue) =>
        val cell = row.createCell(cellIndex)
        val (key, cellValue) = keyValue
        key match {
          case key if key.startsWith(STATIC_TEXT_KEY) => cellValue match {
            case CellString(string: String) => cell.setCellValue(string)
            case _ => cell.setCellType(CellType.BLANK)
          }

          case key if key.startsWith(STATIC_NUMBER_KEY) => cellValue match {
            case CellDouble(double: Double) => cell.setCellValue(double)
            case _ => cell.setCellType(CellType.BLANK)
          }

          case key if key.startsWith(ExportModelUtils.SUBSTITUTION_KEY) || key.startsWith(ExportModelUtils.REPEATED_FIELD_KEY) =>
            if (substituteKey) cell.setCellValue(key)
            else cellValue match {
              case CellString(string: String) => cell.setCellValue(string)
              case CellDouble(double: Double) => cell.setCellValue(double)
              case CellBoolean(boolean: Boolean) => cell.setCellValue(boolean)
              case CellDate(date: Date) => cell.setCellValue(date)
              case _ => cell.setCellType(CellType.BLANK)
            }

          case _ => cell.setCellType(CellType.BLANK)
        }

        cellIndex + 1
    }
  }

  /*
   * Test models to test substitution and repeated templates
   */
  case class StaticTemplateModel(stringField: Option[String], longField: Option[Long], doubleField: Option[Double])

  case class RepTemplateModel(stringField: Option[String], longField: Option[Long], doubleField: Option[Double])

}
