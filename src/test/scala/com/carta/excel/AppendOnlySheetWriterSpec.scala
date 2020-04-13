package com.carta.excel

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.util.Date

import com.carta.excel.ExcelTestHelpers.{PictureData, PicturePosition, RepTemplateModel, StaticTemplateModel, generateTestWorkbook, getModelSeq}
import com.carta.exscalabur.{DataCell, DataRow}
import com.carta.yaml.{DataType, ExcelType, YamlEntry}
import org.apache.poi.ss.usermodel.{CellType, Workbook}
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.{XSSFCell, XSSFRow, XSSFWorkbook}
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers
import resource.managed

import scala.collection.JavaConverters._
import scala.collection.mutable
import scala.util.Random


class AppendOnlySheetWriterSpec extends AnyFlatSpec with Matchers {

  "SheetWriter" should "copy cellStyles from template workbook to output workbook" in {
    val sheetName = "cellStyles"
    val templateWorkbook = new XSSFWorkbook()
    val outputWorkbook = new SXSSFWorkbook(100)

    val templateStream = new ByteArrayOutputStream()
    val actualStream = new ByteArrayOutputStream()

    val sheet = templateWorkbook.createSheet(sheetName)
    val row = sheet.createRow(0)
    (0 until 10).foreach(i => getRandomCell(templateWorkbook, row, i))
    templateWorkbook.write(templateStream)

    val sheetWriter = AppendOnlySheetWriter(
      templateSheet = templateWorkbook.getSheet(sheetName),
      outputWorkbook = outputWorkbook,
      schema = Map.empty,
      cellStyleCache = mutable.Map.empty
    )
    sheetWriter.writeData(Iterator.empty)

    outputWorkbook.write(actualStream)

    ExcelTestHelpers.assertEqualsWorkbooks(
      new ByteArrayInputStream(templateStream.toByteArray),
      new ByteArrayInputStream(actualStream.toByteArray),
    )
  }

  "SheetWriter" should "not put image if template has no images" in {
    val sheetName = "noImages"
    val templateStream = new ByteArrayOutputStream()
    val actualStream = new ByteArrayOutputStream()
    val templateWorkbook = new XSSFWorkbook()
    val outputWorkbook = new SXSSFWorkbook(100)
    val sheet = templateWorkbook.createSheet(sheetName)
    sheet.createRow(0)
    templateWorkbook.write(templateStream)

    val sheetWriter = AppendOnlySheetWriter(
      templateSheet = templateWorkbook.getSheet(sheetName),
      outputWorkbook = outputWorkbook,
      schema = Map.empty,
      cellStyleCache = mutable.Map.empty
    )
    sheetWriter.writeData(Iterator.empty)

    outputWorkbook.write(actualStream)
    val resultingWorkbook = new XSSFWorkbook(new ByteArrayInputStream(actualStream.toByteArray))
    assert(resultingWorkbook.getAllPictures.isEmpty)
  }

  "SheetWriter" should "put image if template has image" in {
    val sheetName = "image"
    val templateStream = new ByteArrayOutputStream()
    val actualStream = new ByteArrayOutputStream()

    val templateWorkbook = new XSSFWorkbook()
    val outputWorkbook = new SXSSFWorkbook(100)
    val sheet = templateWorkbook.createSheet(sheetName)
    sheet.createRow(0)

    val picType = Workbook.PICTURE_TYPE_PNG
    val picData = PictureData(
      imgSrc = "/img/carta.png",
      imgType = picType,
      posn = PicturePosition(
        col1 = 2,
        row1 = 2
      )
    )

    ExcelTestHelpers.putPicture(sheet, picData)
    templateWorkbook.write(templateStream)

    val sheetWriter = AppendOnlySheetWriter(
      templateSheet = templateWorkbook.getSheet(sheetName),
      outputWorkbook = outputWorkbook,
      schema = Map.empty,
      cellStyleCache = mutable.Map.empty
    )
    sheetWriter.writeData(Iterator.empty)

    outputWorkbook.write(actualStream)

    val resultingWorkbook = new XSSFWorkbook(new ByteArrayInputStream(actualStream.toByteArray))
    assert(resultingWorkbook.getAllPictures.size() == 1)
    val resultingImg = resultingWorkbook.getAllPictures.get(0)
    assert(resultingImg.getPictureType == picType)
  }

  "SheetWriter" should "put multiple images if template has many images" in {
    val sheetName = "image"
    val templateStream = new ByteArrayOutputStream()
    val actualStream = new ByteArrayOutputStream()
    val templateWorkbook = new XSSFWorkbook()
    val outputWorkbook = new SXSSFWorkbook(100)
    val sheet = templateWorkbook.createSheet(sheetName)
    sheet.createRow(0)
    val picType = Workbook.PICTURE_TYPE_PNG

    val picData1 = PictureData(
      imgSrc = "/img/carta.png",
      imgType = picType,
      posn = PicturePosition(
        col1 = 2,
        row1 = 2
      )
    )

    val picData2 = PictureData(
      imgSrc = picData1.imgSrc,
      imgType = picData1.imgType,
      posn = PicturePosition(
        col1 = 5,
        col2 = 6,
        row1 = 3,
        row2 = 8
      )
    )

    ExcelTestHelpers.putPicture(sheet, picData1)
    ExcelTestHelpers.putPicture(sheet, picData2)

    templateWorkbook.write(templateStream)

    val sheetWriter = AppendOnlySheetWriter(
      templateSheet = templateWorkbook.getSheet(sheetName),
      outputWorkbook = outputWorkbook,
      schema = Map.empty,
      cellStyleCache = mutable.Map.empty
    )
    sheetWriter.writeData(Iterator.empty)
    outputWorkbook.write(actualStream)

    val resultingWorkbook = new XSSFWorkbook(new ByteArrayInputStream(actualStream.toByteArray))
    assert(resultingWorkbook.getAllPictures.size() == 2)

    resultingWorkbook.getAllPictures.asScala foreach { picData =>
      assert(picData.getPictureType == picType)
    }
  }

  "SheetWriter" should "not create a new cellStyle for each cell in the output Workbook" in {
    val sheetName = "singleCellStyle"
    val templateStream = new ByteArrayOutputStream()
    val actualStream = new ByteArrayOutputStream()
    val templateWorkbook = new XSSFWorkbook()
    val outputWorkbook = new SXSSFWorkbook(100)

    val sheet = templateWorkbook.createSheet(sheetName)
    val row = sheet.createRow(0)

    // Only add 1 cell style to the template
    val cellStyle = ExcelTestHelpers.randomCellStyle(templateWorkbook)
    (0 until 10).foreach { i =>
      val cell = row.createCell(i)
      cell.setCellStyle(cellStyle)
      cell.setCellType(CellType.STRING)
      cell.setCellValue(Random.nextString(5))
    }

    templateWorkbook.write(templateStream)

    AppendOnlySheetWriter(templateWorkbook.getSheet(sheetName), outputWorkbook, Map.empty, mutable.Map.empty)
      .writeData(Iterator.empty)
    outputWorkbook.write(actualStream)

    // Verify that a new cellStyle was not created for every cell
    // Should be 2 because one default cellStyle and one cellStyle added from the template workbook

    managed(toWorkbook(actualStream)).acquireAndGet { result =>
      result.getNumCellStyles shouldEqual 2
    }
  }

  "SheetWriter" should "copy column widths from template workbook to output workbook" in {
    val sheetName = "columnWidths"
    val templateStream = new ByteArrayOutputStream()
    val actualStream = new ByteArrayOutputStream()
    val templateWorkbook = new XSSFWorkbook()
    val outputWorkbook = new SXSSFWorkbook(100)

    val sheet = templateWorkbook.createSheet(sheetName)
    val row = sheet.createRow(0)

    // Column width is set at the sheet level
    (0 until 10).foreach { i =>
      row.createCell(i)
      sheet.setColumnWidth(i, (i + 1) * 100)
    }

    templateWorkbook.write(templateStream)

    AppendOnlySheetWriter(templateWorkbook.getSheet(sheetName), outputWorkbook, Map.empty, mutable.Map.empty)
      .writeData(Iterator.empty)
    outputWorkbook.write(actualStream)

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

  "SheetWriter" should "do substitution from keys in template workbook to output workbook (happy path)" in {
    val sheetName = "substitutionHappyPath"
    val outputWorkbook = new SXSSFWorkbook(100)
    val templateStream = new ByteArrayOutputStream()
    val expectedStream = new ByteArrayOutputStream()
    val actualStream = new ByteArrayOutputStream()

    val model = StaticTemplateModel(Some("text"), Some(1234), Some(1234))
    val modelSchema = ExcelTestHelpers.getModelSchema(model)
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

    val templateWorkbook = generateTestWorkbook(sheetName, templateRows, insertRowModelKeys = true)
    templateWorkbook.write(templateStream)
    val expectedWorkbook = generateTestWorkbook(sheetName, templateRows)
    expectedWorkbook.write(expectedStream)

    val sheetWriter = AppendOnlySheetWriter(
      templateWorkbook.getSheet(sheetName),
      outputWorkbook,
      modelSchema,
      mutable.Map.empty
    )

    sheetWriter.writeData(Iterator((
      Vector(
        DataCell(s"string_field", "text"),
        DataCell(s"long_field", 1234),
        DataCell(s"double_field", 1234.0)
      ),
      Vector.empty
    )))

    outputWorkbook.write(actualStream)

    val expectedInputStream = new ByteArrayInputStream(expectedStream.toByteArray)
    val actualInputStream = new ByteArrayInputStream(actualStream.toByteArray)

    ExcelTestHelpers.assertEqualsWorkbooks(expectedInputStream, actualInputStream)
  }

  "SheetWriter" should "do substitution from keys in template workbook despite missing values in template" in {
    val sheetName = "substitutionMissingValues"
    val outputWorkbook = new SXSSFWorkbook(100)
    val templateStream = new ByteArrayOutputStream()
    val expectedStream = new ByteArrayOutputStream()
    val actualStream = new ByteArrayOutputStream()

    val model = StaticTemplateModel(Some("text"), Some(1234), Some(1234))
    val modelSchema = ExcelTestHelpers.getModelSchema(model)
    val modelSeq = getModelSeq(model)

    // Build template
    val templateRows = Seq(
      // Row with static text, empty cell, and number
      Seq(
        (ExcelTestHelpers.STATIC_DOUBLE_KEY, CellDouble(12.34)),
        (ExcelTestHelpers.STATIC_BOOL_KEY, CellBoolean(false)),
      ),

      // Row with a substitution key missing
      modelSeq.drop(1)
    )

    val templateWorkbook: XSSFWorkbook = generateTestWorkbook(sheetName, templateRows, insertRowModelKeys = true)
    templateWorkbook.write(templateStream)
    val expectedWorkbook: XSSFWorkbook = generateTestWorkbook(sheetName, templateRows)
    expectedWorkbook.write(expectedStream)

    val data = Iterator((
      Vector(
        DataCell(s"string_field", "text"),
        DataCell(s"long_field", 1234),
        DataCell(s"double_field", 1234.0)
      ),
      Vector.empty
    ))

    val sheetWriter = AppendOnlySheetWriter(
      templateWorkbook.getSheet(sheetName),
      outputWorkbook,
      modelSchema,
      mutable.Map.empty
    )
    sheetWriter.writeData(data)
    outputWorkbook.write(actualStream)

    val expectedInputStream = new ByteArrayInputStream(expectedStream.toByteArray)
    val actualInputStream = new ByteArrayInputStream(actualStream.toByteArray)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedInputStream, actualInputStream)
  }

  "SheetWriter" should "do substitution from keys in template workbook despite Blank values in model" in {
    val sheetName = "substitutionExtraKeys"
    val outputWorkbook = new SXSSFWorkbook(100)
    val templateStream = new ByteArrayOutputStream()
    val expectedStream = new ByteArrayOutputStream()
    val actualStream = new ByteArrayOutputStream()

    val model = StaticTemplateModel(Some("text"), None, Some(1234))
    val modelSchema = ExcelTestHelpers.getModelSchema(model)
    val modelSeq = getModelSeq(model)

    val templateRows = Seq(
      // Row with static text, empty cell, and number
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("text")),
        ("empty", CellBlank()),
        (ExcelTestHelpers.STATIC_DATE_KEY, CellDate(new Date(2019, 12, 13)))
      ),

      // Row with substitution keys
      modelSeq
    )
    val expectedRows = Seq(
      // Row with static text, empty cell, and number
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("text")),
        ("empty", CellBlank()),
        (ExcelTestHelpers.STATIC_DATE_KEY, CellDate(new Date(2019, 12, 13))),
      ),

      // Expect row with 1 None value missing
      modelSeq
    )

    val templateWorkbook: XSSFWorkbook = ExcelTestHelpers.generateTestWorkbook(sheetName, templateRows, insertRowModelKeys = true)
    templateWorkbook.write(templateStream)
    val expectedWorkbook: XSSFWorkbook = ExcelTestHelpers.generateTestWorkbook(sheetName, expectedRows)
    expectedWorkbook.write(expectedStream)

    val data = Iterator((
      Vector(
        DataCell(s"string_field", "text"),
        DataCell(s"double_field", 1234.0)
      ),
      Vector.empty
    ))
    val sheetWriter = AppendOnlySheetWriter(templateWorkbook.getSheet(sheetName), outputWorkbook, modelSchema, mutable.Map.empty)
    sheetWriter.writeData(data)
    outputWorkbook.write(actualStream)

    val expectedInputStream = new ByteArrayInputStream(expectedStream.toByteArray)
    val actualInputStream = new ByteArrayInputStream(actualStream.toByteArray)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedInputStream, actualInputStream)
  }

  "SheetWriter" should "load template for repeated rows and then insert set of repeated rows" in {
    val templateName = "substitutionExtraKeys"
    val outputWorkbook = new SXSSFWorkbook(100)
    val templateStream = new ByteArrayOutputStream()
    val expectedStream = new ByteArrayOutputStream()
    val actualStream = new ByteArrayOutputStream()

    val repeatedModels = Seq(
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
    )
    val schema = Map(
      "$REP.string_field" -> YamlEntry(DataType.string, ExcelType.string),
      "$REP.long_field" -> YamlEntry(DataType.long, ExcelType.number),
      "$REP.double_field" -> YamlEntry(DataType.double, ExcelType.number),
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

    val templateBook = ExcelTestHelpers.generateTestWorkbook(templateName, templateRows, insertRowModelKeys = true)
    templateBook.write(templateStream)
    ExcelTestHelpers.generateTestWorkbook(templateName, expectedRows)
      .write(expectedStream)

    val data = Iterator((
      Vector.empty,
      Vector(
        DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
        DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
        DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
        DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
        DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
      )
    ))

    val sheetWriter = AppendOnlySheetWriter(templateBook.getSheet(templateName), outputWorkbook, schema, mutable.Map.empty)
    sheetWriter.writeData(data)

    outputWorkbook.write(actualStream)

    val expectedInputStream = new ByteArrayInputStream(expectedStream.toByteArray)
    val actualInputStream = new ByteArrayInputStream(actualStream.toByteArray)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedInputStream, actualInputStream)
  }

  "SheetWriter" should "copy repeated rows followed by static rows without overwriting" in {
    val templateName = "copyRepeatedThenStatic"
    val outputWorkbook = new SXSSFWorkbook(100)
    val templateStream = new ByteArrayOutputStream()
    val expectedStream = new ByteArrayOutputStream()
    val actualStream = new ByteArrayOutputStream()

    val repeatedModels = Seq(
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
    )

    val lowerStaticModels = Seq(
      StaticTemplateModel(Some("Test Word"), Some(5432), None)
    )
    val repeatedModelsSeq = repeatedModels.map(getModelSeq)
    val lowerStaticModelsSeq = lowerStaticModels.map(getModelSeq)

    val schema = Map(
      "$REP.string_field" -> YamlEntry(DataType.string, ExcelType.string),
      "$REP.long_field" -> YamlEntry(DataType.long, ExcelType.number),
      "$REP.double_field" -> YamlEntry(DataType.double, ExcelType.number),
      "$KEY.string_field" -> YamlEntry(DataType.string, ExcelType.string),
      "$KEY.long_field" -> YamlEntry(DataType.long, ExcelType.number),
    )

    val templateRows: Seq[Iterable[(String, CellValue)]] = Seq(
      // Row with static text, empty cell, and number
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("static_text")),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1337)),
      ),
      repeatedModelsSeq.head,
      lowerStaticModelsSeq.head
    )

    val expectedRows = Seq(
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("static_text")),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1337))
      ),
    ) ++ repeatedModelsSeq ++ lowerStaticModelsSeq

    val templateBook = ExcelTestHelpers.generateTestWorkbook(templateName, templateRows, insertRowModelKeys = true)
    templateBook.write(templateStream)
    ExcelTestHelpers.generateTestWorkbook(templateName, expectedRows)
      .write(expectedStream)

    val data = Iterator((
      Vector(
        DataCell("string_field", "Test Word"),
        DataCell("long_field", 5432),
      ),
      Vector(
        DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
        DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
        DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
        DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
        DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
      )
    ))
    val sheetWriter = AppendOnlySheetWriter(templateBook.getSheet(templateName), outputWorkbook, schema, mutable.Map.empty)
    sheetWriter.writeData(data)
    outputWorkbook.write(actualStream)

    val expectedInputStream = new ByteArrayInputStream(expectedStream.toByteArray)
    val actualInputStream = new ByteArrayInputStream(actualStream.toByteArray)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedInputStream, actualInputStream)
  }

  "SheetWriter" should "copy data in chunks" in {
    val templateName = "copyRepeatedThenStatic"
    val outputWorkbook = new SXSSFWorkbook(100)
    val templateStream = new ByteArrayOutputStream()
    val expectedStream = new ByteArrayOutputStream()
    val actualStream = new ByteArrayOutputStream()

    val repeatedModels = Seq(
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
    )

    val lowerStaticModels = Seq(
      StaticTemplateModel(Some("Test Word"), Some(5432), None)
    )
    val repeatedModelsSeq = repeatedModels.map(getModelSeq)
    val lowerStaticModelsSeq = lowerStaticModels.map(getModelSeq)

    val schema = Map(
      "$REP.string_field" -> YamlEntry(DataType.string, ExcelType.string),
      "$REP.long_field" -> YamlEntry(DataType.long, ExcelType.number),
      "$REP.double_field" -> YamlEntry(DataType.double, ExcelType.number),
      "$KEY.string_field" -> YamlEntry(DataType.string, ExcelType.string),
      "$KEY.long_field" -> YamlEntry(DataType.long, ExcelType.number),
    )

    val templateRows: Seq[Iterable[(String, CellValue)]] = Seq(
      // Row with static text, empty cell, and number
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("static text")),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1235)),
      ),
      repeatedModelsSeq.head,
      lowerStaticModelsSeq.head
    )

    val expectedRows = Seq(
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("static text")),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1235))
      ),
    ) ++ repeatedModelsSeq ++ lowerStaticModelsSeq

    val templateBook = ExcelTestHelpers.generateTestWorkbook(templateName, templateRows, insertRowModelKeys = true)
    templateBook.write(templateStream)
    ExcelTestHelpers.generateTestWorkbook(templateName, expectedRows)
      .write(expectedStream)

    val data = Iterator((
      Vector.empty,
      Vector(
        DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
        DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build()
      )), (
      Vector.empty,
      Vector(
        DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
        DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
        DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
      )), (
      Vector(
        DataCell("string_field", "Test Word"),
        DataCell("long_field", 5432),
      ),
      Vector.empty
    ))


    val sheetWriter = AppendOnlySheetWriter(templateBook.getSheet(templateName), outputWorkbook, schema, mutable.Map.empty)
    sheetWriter.writeData(data)
    outputWorkbook.write(actualStream)

    val expectedInputStream = new ByteArrayInputStream(expectedStream.toByteArray)
    val actualInputStream = new ByteArrayInputStream(actualStream.toByteArray)

    ExcelTestHelpers.assertEqualsWorkbooks(expectedInputStream, actualInputStream)
  }

  "SheetWriter" should "load template with static text between repeated rows" in {
    val templateName = "staticText"
    val outputWorkbook = new SXSSFWorkbook(100)
    val templateStream = new ByteArrayOutputStream()
    val expectedStream = new ByteArrayOutputStream()
    val actualStream = new ByteArrayOutputStream()

    val schema = Map(
      "$REP.string_field" -> YamlEntry(DataType.string, ExcelType.string),
      "$REP.long_field" -> YamlEntry(DataType.long, ExcelType.number),
      "$REP.double_field" -> YamlEntry(DataType.double, ExcelType.number),

      "$REP.string_field2" -> YamlEntry(DataType.string, ExcelType.string),
      "$REP.long_field2" -> YamlEntry(DataType.long, ExcelType.number),
      "$REP.double_field2" -> YamlEntry(DataType.double, ExcelType.number),
    )

    val repeatedModels = List(
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
    )
    val repeatedModelsSeq1 = repeatedModels.map(getModelSeq)
    val repeatedModelsSeq2 = repeatedModels.map(getModelSeq).map(_.map { case (key, cellValue) => (f"${key}2", cellValue) })

    val templateRows: Seq[Iterable[(String, CellValue)]] = Seq(
      // Row with static text, empty cell, and number
      repeatedModelsSeq1.head,
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("static_text")),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1337)),
      ),
      repeatedModelsSeq2.head
    )

    val expectedRows = repeatedModelsSeq1 ++ Seq(
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("static_text")),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1337))
      )
    ) ++ repeatedModelsSeq2

    val templateBook = ExcelTestHelpers.generateTestWorkbook(templateName, templateRows, insertRowModelKeys = true)
    templateBook.write(templateStream)
    ExcelTestHelpers.generateTestWorkbook(templateName, expectedRows)
      .write(expectedStream)

    val data1 = Vector(
      DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
      DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
      DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
      DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
      DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
    )

    val data2 = Vector(
      DataRow.Builder().addCell("string_field2", "text").addCell("long_field2", 1234).addCell("double_field2", 1234.0).build(),
      DataRow.Builder().addCell("string_field2", "text").addCell("long_field2", 1234).addCell("double_field2", 1234.0).build(),
      DataRow.Builder().addCell("string_field2", "text").addCell("long_field2", 1234).addCell("double_field2", 1234.0).build(),
      DataRow.Builder().addCell("string_field2", "text").addCell("long_field2", 1234).addCell("double_field2", 1234.0).build(),
      DataRow.Builder().addCell("string_field2", "text").addCell("long_field2", 1234).addCell("double_field2", 1234.0).build(),
    )

    val data = Iterator((Vector.empty, data1), (Vector.empty, data2))
    val sheetWriter = AppendOnlySheetWriter(templateBook.getSheet(templateName), outputWorkbook, schema, mutable.Map.empty)
    sheetWriter.writeData(data)
    outputWorkbook.write(actualStream)

    val expectedInputStream = new ByteArrayInputStream(expectedStream.toByteArray)
    val actualInputStream = new ByteArrayInputStream(actualStream.toByteArray)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedInputStream, actualInputStream)
  }

  "SheetWriter" should "load template with static text between substitution rows" in {
    val templateName = "staticText"
    val outputWorkbook = new SXSSFWorkbook(100)
    val templateStream = new ByteArrayOutputStream()
    val expectedStream = new ByteArrayOutputStream()
    val actualStream = new ByteArrayOutputStream()

    val schema = Map(
      "$KEY.string_field" -> YamlEntry(DataType.string, ExcelType.string),
      "$KEY.long_field" -> YamlEntry(DataType.long, ExcelType.number),
      "$KEY.double_field" -> YamlEntry(DataType.double, ExcelType.number),

      "$KEY.string_field2" -> YamlEntry(DataType.string, ExcelType.string),
      "$KEY.long_field2" -> YamlEntry(DataType.long, ExcelType.number),
      "$KEY.double_field2" -> YamlEntry(DataType.double, ExcelType.number),
    )

    val staticModel = StaticTemplateModel(Some("text"), Some(1233), Some(1234))

    val staticModel2 = StaticTemplateModel(Some("text2"), Some(1235), Some(1236))

    val modelSeq1 = getModelSeq(staticModel)
    val modelSeq2 = getModelSeq(staticModel2).map { case (key, cellValue) => (f"${key}2", cellValue) }

    val templateRows = Seq(
      modelSeq1,
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("static_text")),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1337)),
      ),
      modelSeq2
    )

    val expectedRows: Seq[Iterable[(String, CellValue)]] = Seq(
      modelSeq1,
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("static_text")),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1337))
      ),
      modelSeq2
    )

    val templateBook = ExcelTestHelpers.generateTestWorkbook(templateName, templateRows, insertRowModelKeys = true)
    templateBook.write(templateStream)
    ExcelTestHelpers.generateTestWorkbook(templateName, expectedRows)
      .write(expectedStream)

    val data1 = Vector(
      DataCell("string_field", "text"),
      DataCell("long_field", 1233),
      DataCell("double_field", 1234.0)
    )
    val data2 = Vector(
      DataCell("string_field2", "text2"),
      DataCell("long_field2", 1235),
      DataCell("double_field2", 1236.0)
    )

    val data = Iterator((data1, Vector.empty), (data2, Vector.empty))
    val sheetWriter = AppendOnlySheetWriter(templateBook.getSheet(templateName), outputWorkbook, schema, mutable.Map.empty)
    sheetWriter.writeData(data)
    outputWorkbook.write(actualStream)

    val expectedInputStream = new ByteArrayInputStream(expectedStream.toByteArray)
    val actualInputStream = new ByteArrayInputStream(actualStream.toByteArray)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedInputStream, actualInputStream)
  }

  "SheetWriter" should "load template with static text between repeated then substitution row" in {
    val templateName = "staticText"
    val outputWorkbook = new SXSSFWorkbook(100)
    val templateStream = new ByteArrayOutputStream()
    val expectedStream = new ByteArrayOutputStream()
    val actualStream = new ByteArrayOutputStream()

    val schema = Map(
      "$REP.string_field" -> YamlEntry(DataType.string, ExcelType.string),
      "$REP.long_field" -> YamlEntry(DataType.long, ExcelType.number),
      "$REP.double_field" -> YamlEntry(DataType.double, ExcelType.number),

      "$KEY.string_field" -> YamlEntry(DataType.string, ExcelType.string),
      "$KEY.long_field" -> YamlEntry(DataType.long, ExcelType.number),
      "$KEY.double_field" -> YamlEntry(DataType.double, ExcelType.number),
    )

    val repeatedModels = List(
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
    )
    val staticModel = StaticTemplateModel(Some("static text"), Some(12), Some(12.4))


    val repeatedModelsSeq1 = repeatedModels.map(getModelSeq)
    val staticModelSeq = getModelSeq(staticModel)

    val templateRows: Seq[Iterable[(String, CellValue)]] = Seq(
      // Row with static text, empty cell, and number
      repeatedModelsSeq1.head,
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("static_text")),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1337)),
      ),
      staticModelSeq
    )

    val expectedRows = repeatedModelsSeq1 ++ Seq(
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("static_text")),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1337))
      ),
      staticModelSeq
    )

    val templateBook = ExcelTestHelpers.generateTestWorkbook(templateName, templateRows, insertRowModelKeys = true)
    templateBook.write(templateStream)
    ExcelTestHelpers.generateTestWorkbook(templateName, expectedRows)
      .write(expectedStream)

    val repeatedData = Vector(
      DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
      DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
      DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
      DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
      DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
    )

    val staticData = Vector(
      DataCell("string_field", "static text"),
      DataCell("long_field", 12),
      DataCell("double_field", 12.4)
    )

    val data = Iterator((Vector.empty, repeatedData), (staticData, Vector.empty))
    val sheetWriter = AppendOnlySheetWriter(templateBook.getSheet(templateName), outputWorkbook, schema, mutable.Map.empty)
    sheetWriter.writeData(data)
    outputWorkbook.write(actualStream)

    val expectedInputStream = new ByteArrayInputStream(expectedStream.toByteArray)
    val actualInputStream = new ByteArrayInputStream(actualStream.toByteArray)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedInputStream, actualInputStream)
  }

  "SheetWriter" should "load template with static text between substitution then repeated row" in {
    val templateName = "staticText"
    val outputWorkbook = new SXSSFWorkbook(100)
    val templateStream = new ByteArrayOutputStream()
    val expectedStream = new ByteArrayOutputStream()
    val actualStream = new ByteArrayOutputStream()

    val schema = Map(
      "$REP.string_field" -> YamlEntry(DataType.string, ExcelType.string),
      "$REP.long_field" -> YamlEntry(DataType.long, ExcelType.number),
      "$REP.double_field" -> YamlEntry(DataType.double, ExcelType.number),

      "$KEY.string_field" -> YamlEntry(DataType.string, ExcelType.string),
      "$KEY.long_field" -> YamlEntry(DataType.long, ExcelType.number),
      "$KEY.double_field" -> YamlEntry(DataType.double, ExcelType.number),
    )

    val repeatedModels = List(
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
      RepTemplateModel(Some("text"), Some(1234), Some(1234)),
    )
    val staticModel = StaticTemplateModel(Some("static text"), Some(12), Some(12.4))


    val repeatedModelsSeq1 = repeatedModels.map(getModelSeq)
    val staticModelSeq = getModelSeq(staticModel)

    val templateRows: Seq[Iterable[(String, CellValue)]] = Seq(
      // Row with static text, empty cell, and number
      staticModelSeq,
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("static_text")),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1337)),
      ),
      repeatedModelsSeq1.head,
    )

    val expectedRows = Seq(
      staticModelSeq,
      Seq(
        (ExcelTestHelpers.STATIC_TEXT_KEY, CellString("static_text")),
        (ExcelTestHelpers.STATIC_NUMBER_KEY, CellDouble(1337))
      ),
    ) ++ repeatedModelsSeq1

    val templateBook = ExcelTestHelpers.generateTestWorkbook(templateName, templateRows, insertRowModelKeys = true)
    templateBook.write(templateStream)
    ExcelTestHelpers.generateTestWorkbook(templateName, expectedRows)
      .write(expectedStream)

    val repeatedData = Vector(
      DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
      DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
      DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
      DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
      DataRow.Builder().addCell("string_field", "text").addCell("long_field", 1234).addCell("double_field", 1234.0).build(),
    )

    val staticData = Vector(
      DataCell("string_field", "static text"),
      DataCell("long_field", 12),
      DataCell("double_field", 12.4)
    )

    val data = Iterator((staticData, Vector.empty), (Vector.empty, repeatedData))
    val sheetWriter = AppendOnlySheetWriter(templateBook.getSheet(templateName), outputWorkbook, schema, mutable.Map.empty)
    sheetWriter.writeData(data)
    outputWorkbook.write(actualStream)

    val expectedInputStream = new ByteArrayInputStream(expectedStream.toByteArray)
    val actualInputStream = new ByteArrayInputStream(actualStream.toByteArray)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedInputStream, actualInputStream)
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
