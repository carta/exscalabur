package com.carta.excel

import java.io.{File, FileInputStream, FileOutputStream}

import com.carta.exscalabur.{DataRow, Exscalabur}
import com.carta.yaml.{KeyType, YamlCellType, YamlEntry}
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

class ExscalaburSpec extends AnyFlatSpec with Matchers {

  "Exscalabur" should "write excel file to disk" in {
    val filePath = File.createTempFile("ExscalaburSpec", ".xlsx").getAbsolutePath
    val staticData = DataRow.Builder()
      .addCell("col1", "col1Value")
      .addCell("col2", 1.45)
      .build()
      .getCells

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
      "$KEY.col1" -> YamlEntry(KeyType.single, YamlCellType.string, YamlCellType.string),
      "$KEY.col2" -> YamlEntry(KeyType.single, YamlCellType.double, YamlCellType.double),
      "$REP.col1" -> YamlEntry(KeyType.repeated, YamlCellType.string, YamlCellType.string),
      "$REP.col2" -> YamlEntry(KeyType.repeated, YamlCellType.double, YamlCellType.double),
    )

    val exscalabur = Exscalabur(
      List(getClass.getResource("/excel/templates/writerSpec.xlsx").getFile), filePath, schema, 100
    )

    exscalabur.getAppendOnlySheetWriter("Sheet1").writeData(staticData, repeatedData)
    exscalabur.writeOut()

    val expectedWorkbookStream = getClass.getResourceAsStream("/excel/expected/writerSpec.xlsx")
    val actualWorkbookStream = new FileInputStream(filePath)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedWorkbookStream, actualWorkbookStream)
  }

  "Exscalabur" should "write excel file to stream" in {

    val filePath = File.createTempFile("writerSpec", ".xlsx").getAbsolutePath
    val fileStream = new FileOutputStream(filePath)

    val staticData = DataRow.Builder()
      .addCell("col1", "col1Value")
      .addCell("col2", 1.45)
      .build()
      .getCells
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
      "$KEY.col1" -> YamlEntry(KeyType.single, YamlCellType.string, YamlCellType.string),
      "$KEY.col2" -> YamlEntry(KeyType.single, YamlCellType.double, YamlCellType.double),
      "$REP.col1" -> YamlEntry(KeyType.repeated, YamlCellType.string, YamlCellType.string),
      "$REP.col2" -> YamlEntry(KeyType.repeated, YamlCellType.double, YamlCellType.double),
    )

    val templatePath = getClass.getResource("/excel/templates/writerSpec.xlsx").getFile

    val exscalabur = Exscalabur(
      List(templatePath), filePath, schema, 100
    )

    exscalabur.getAppendOnlySheetWriter("Sheet1")
      .writeData(staticData, repeatedData)
    exscalabur.writeOut()

    val expectedWorkbookStream = getClass.getResourceAsStream("/excel/expected/writerSpec.xlsx")
    fileStream.close()
    val actualWorkbookStream = new FileInputStream(filePath)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedWorkbookStream, actualWorkbookStream)
  }

  "Exscalabur" should "keep blank rows when copying sequential repeated rows" in {
    val filePath = File.createTempFile("sepRepRows", ".xlsx").getAbsolutePath

    val animals = List("bear", "eagle", "elephant", "bird", "snake", "pig", "dog", "cat", "penguin", "anteater");
    val weight = List(1, 2, 3, 4, 5, 6, 7, 8, 9, 10);

    val repeatedData = (0 until 10).map(i => {
      DataRow.Builder().addCell("animal", animals(i)).build()
    }) ++ (0 until 10).map { i =>
      DataRow.Builder().addCell("weight", weight(i)).build()
    }

    val schema = Map(
      "$REP.animal" -> YamlEntry(KeyType.repeated, YamlCellType.string, YamlCellType.string),
      "$REP.weight" -> YamlEntry(KeyType.repeated, YamlCellType.long, YamlCellType.double),
    )

    val templatePath = getClass.getResource("/excel/templates/seqRepRows.xlsx").getFile

    val exscalabur = Exscalabur(
      List(templatePath), filePath, schema, 100
    )

    exscalabur.getAppendOnlySheetWriter("Sheet1")
      .writeData(Vector.empty, repeatedData)
    exscalabur.writeOut()

    val expectedWorkbookStream = getClass.getResourceAsStream("/excel/expected/seqRepRows.xlsx")
    val actualWorkbookStream = new FileInputStream(filePath)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedWorkbookStream, actualWorkbookStream)
  }

  "Exscalabur" should "write data in parts" in {
    val filePath = File.createTempFile("sepRepRows", ".xlsx").getAbsolutePath

    val animals = List("bear", "eagle", "elephant", "bird", "snake", "pig", "dog", "cat", "penguin", "anteater");
    val weight = List(1, 2, 3, 4, 5, 6, 7, 8, 9, 10);

    val repeatedData1 = animals.take(5)
      .map(animal => DataRow.Builder().addCell("animal", animal).build())

    val repeatedData2 = animals.slice(5, 10)
      .map(animal => DataRow.Builder().addCell("animal", animal).build())

    val repeatedData3 = weight.map(weight => DataRow.Builder().addCell("weight", weight).build())

    val schema = Map(
      "$REP.animal" -> YamlEntry(KeyType.repeated, YamlCellType.string, YamlCellType.string),
      "$REP.weight" -> YamlEntry(KeyType.repeated, YamlCellType.long, YamlCellType.double),
    )

    val templatePath = getClass.getResource("/excel/templates/seqRepRows.xlsx").getFile

    val exscalabur = Exscalabur(
      List(templatePath), filePath, schema, 100
    )

    val sheetWriter = exscalabur.getAppendOnlySheetWriter("Sheet1")
    sheetWriter.writeData(Vector.empty, repeatedData1, hasMoreDataForRow = true)
    sheetWriter.writeData(Vector.empty, repeatedData2)
    sheetWriter.writeData(Vector.empty, repeatedData3)
    exscalabur.writeOut()

    val expectedWorkbookStream = getClass.getResourceAsStream("/excel/expected/seqRepRows.xlsx")
    val actualWorkbookStream = new FileInputStream(filePath)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedWorkbookStream, actualWorkbookStream)
  }
}
