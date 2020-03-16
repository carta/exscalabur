package com.carta.excel

import java.io.{File, FileInputStream, FileOutputStream}

import com.carta.exscalabur.DataRow
import com.carta.yaml.{CellType, KeyType, YamlEntry}
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

class WriterSpec extends AnyFlatSpec with Matchers {

  "Writer" should "write excel file to disk" in {

    val filePath = File.createTempFile("writerSpec", ".xlsx").getAbsolutePath

    val staticData = DataRow.Builder()
      .addCell("col1", "col1Value")
      .addCell("col2", 1.45)
      .build()
    val repeatedData = List(
      DataRow.Builder()
        .addCell("col1", "repeatedString1")
        .addCell("col2", 1.573)
        .build(),
      DataRow.Builder()
        .addCell("col1", "repeatedString2")
        .addCell("col2", 2.5)
        .build(),
    )

    val schema = Map(
      "$KEY.col1" -> YamlEntry(KeyType.single, CellType.string, CellType.string),
      "$KEY.col2" -> YamlEntry(KeyType.single, CellType.double, CellType.double),
      "$REP.col1" -> YamlEntry(KeyType.repeated, CellType.string, CellType.string),
      "$REP.col2" -> YamlEntry(KeyType.repeated, CellType.double, CellType.double),
    )

    val sheetData = SheetData(
      sheetName = "Sheet1",
      templatePath = getClass.getResource("/excel/templates/writerSpec.xlsx").getFile,
      staticData = staticData.getCells,
      repeatedData = repeatedData,
      schema
    )

    new Writer(100).writeExcelFileToDisk(filePath, List(sheetData))

    val expectedWorkbookStream = getClass.getResourceAsStream("/excel/expected/writerSpec.xlsx")
    val actualWorkbookStream = new FileInputStream(filePath)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedWorkbookStream, actualWorkbookStream)
  }

  "Writer" should "write excel file to stream" in {

    val filePath = File.createTempFile("writerSpec", ".xlsx").getAbsolutePath
    val fileStream = new FileOutputStream(filePath)

    val staticData = DataRow.Builder()
      .addCell("col1", "col1Value")
      .addCell("col2", 1.45)
      .build()
    val repeatedData = List(
      DataRow.Builder()
        .addCell("col1", "repeatedString1")
        .addCell("col2", 1.573)
        .build(),
      DataRow.Builder()
        .addCell("col1", "repeatedString2")
        .addCell("col2", 2.5)
        .build(),
    )

    val schema = Map(
      "$KEY.col1" -> YamlEntry(KeyType.single, CellType.string, CellType.string),
      "$KEY.col2" -> YamlEntry(KeyType.single, CellType.double, CellType.double),
      "$REP.col1" -> YamlEntry(KeyType.repeated, CellType.string, CellType.string),
      "$REP.col2" -> YamlEntry(KeyType.repeated, CellType.double, CellType.double),
    )

    val sheetData = SheetData(
      sheetName = "Sheet1",
      templatePath = getClass.getResource("/excel/templates/writerSpec.xlsx").getFile,
      staticData = staticData.getCells,
      repeatedData = repeatedData,
      schema = schema
    )

    new Writer(100).writeExcelToStream(List(sheetData), fileStream)

    val expectedWorkbookStream = getClass.getResourceAsStream("/excel/expected/writerSpec.xlsx")
    fileStream.close()
    val actualWorkbookStream = new FileInputStream(filePath)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedWorkbookStream, actualWorkbookStream)
  }

  "Writer" should "keep blank rows when copying sequential repeated rows" in {
    val filePath = File.createTempFile("sepRepRows", ".xlsx").getAbsolutePath


    val animals = List("bear", "eagle", "elephant", "bird", "snake", "pig", "dog", "cat", "penguin", "anteater");
    val weight = List(1, 2, 3, 4, 5, 6, 7, 8, 9, 10);

    val repeatedData = (0 until 10).map(i => {
      DataRow.Builder().addCell("animal", animals(i)).build()
    })  ++ (0 until 10).map { i =>
      DataRow.Builder().addCell("weight", weight(i)).build()
    }

    val schema = Map(
      "$REP.animal" -> YamlEntry(KeyType.repeated, CellType.string, CellType.string),
      "$REP.weight" -> YamlEntry(KeyType.repeated, CellType.long, CellType.double),
    )

    val sheetData = SheetData(
      sheetName = "Sheet1",
      templatePath = getClass.getResource("/excel/templates/seqRepRows.xlsx").getFile,
      staticData = Vector.empty,
      repeatedData = repeatedData.toList,
      schema
    )

    new Writer(100).writeExcelFileToDisk(filePath, List(sheetData))

    val expectedWorkbookStream = getClass.getResourceAsStream("/excel/expected/seqRepRows.xlsx")
    val actualWorkbookStream = new FileInputStream(filePath)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedWorkbookStream, actualWorkbookStream)
  }
}
