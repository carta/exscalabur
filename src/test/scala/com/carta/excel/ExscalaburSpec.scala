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

import java.io.{FileInputStream, FileOutputStream}

import com.carta.exscalabur.{DataCell, DataRow, Exscalabur}
import com.carta.yaml.{DataType, ExcelType, YamlEntry}
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

class ExscalaburSpec extends AnyFlatSpec with Matchers {

  "Exscalabur" should "write excel file to disk" in {
    val filePath = ExcelTestHelpers.createTempFile("ExscalaburSpec").getAbsolutePath
    val staticData = DataRow()
      .addCell("col1", "col1Value")
      .addCell("col2", 1.45)
      .getCells

    val repeatedData = List(
      DataRow()
        .addCell("col1", "repeatedString1")
        .addCell("col2", 1.573),
      DataRow()
        .addCell("col1", "repeatedString2")
        .addCell("col2", 2.5),
    )

    val schema = Map(
      "$KEY.col1" -> YamlEntry(DataType.string, ExcelType.string),
      "$KEY.col2" -> YamlEntry(DataType.double, ExcelType.number),
      "$REP.col1" -> YamlEntry(DataType.string, ExcelType.string),
      "$REP.col2" -> YamlEntry(DataType.double, ExcelType.number),
    )

    val exscalabur = Exscalabur(
      List(getClass.getResource("/excel/templates/writerSpec.xlsx").getFile), schema, 100
    )

    val sheetWriter = exscalabur.getAppendOnlySheetWriter("Sheet1")
    sheetWriter.writeStaticData(staticData)
    sheetWriter.writeRepeatedData(repeatedData)

    exscalabur.exportToFile(filePath)

    val expectedWorkbookStream = getClass.getResourceAsStream("/excel/expected/writerSpec.xlsx")
    val actualWorkbookStream = new FileInputStream(filePath)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedWorkbookStream, actualWorkbookStream)
  }

  "Exscalabur" should "write excel file to stream" in {
    val filePath = ExcelTestHelpers.createTempFile("writerSpec").getAbsolutePath
    val fileStream = new FileOutputStream(filePath)

    val staticData = DataRow()
      .addCell("col1", "col1Value")
      .addCell("col2", 1.45)

      .getCells
    val repeatedData = List(
      DataRow()
        .addCell("col1", "repeatedString1")
        .addCell("col2", 1.573),
      DataRow()
        .addCell("col1", "repeatedString2")
        .addCell("col2", 2.5),
    )

    val schema = Map(
      "$KEY.col1" -> YamlEntry(DataType.string, ExcelType.string),
      "$KEY.col2" -> YamlEntry(DataType.double, ExcelType.number),
      "$REP.col1" -> YamlEntry(DataType.string, ExcelType.string),
      "$REP.col2" -> YamlEntry(DataType.double, ExcelType.number),
    )

    val templatePath = getClass.getResource("/excel/templates/writerSpec.xlsx").getFile

    val exscalabur = Exscalabur(List(templatePath), schema, 100)

    val sheetWriter = exscalabur.getAppendOnlySheetWriter("Sheet1")

    sheetWriter.writeStaticData(staticData)
    sheetWriter.writeRepeatedData(repeatedData)
    exscalabur.exportToFile(filePath)

    val expectedWorkbookStream = getClass.getResourceAsStream("/excel/expected/writerSpec.xlsx")
    fileStream.close()
    val actualWorkbookStream = new FileInputStream(filePath)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedWorkbookStream, actualWorkbookStream)
  }

  "Exscalabur" should "keep blank rows when copying sequential repeated rows" in {
    val filePath = ExcelTestHelpers.createTempFile("sepRepRows").getAbsolutePath
    val animals = List("bear", "eagle", "elephant", "bird", "snake", "pig", "dog", "cat", "penguin", "anteater");
    val weight = List(1, 2, 3, 4, 5, 6, 7, 8, 9, 10);

    val repeatedData = (0 until 10).map(i => {
      DataRow().addCell("animal", animals(i))
    }) ++ (0 until 10).map { i =>
      DataRow().addCell("weight", weight(i))
    }

    val schema = Map(
      "$REP.animal" -> YamlEntry(DataType.string, ExcelType.string),
      "$REP.weight" -> YamlEntry(DataType.long, ExcelType.number),
    )

    val templatePath = getClass.getResource("/excel/templates/seqRepRows.xlsx").getFile

    val exscalabur = Exscalabur(List(templatePath), schema, 100)

    exscalabur.getAppendOnlySheetWriter("Sheet1")
      .writeRepeatedData(repeatedData)
    exscalabur.exportToFile(filePath)

    val expectedWorkbookStream = getClass.getResourceAsStream("/excel/expected/seqRepRows.xlsx")
    val actualWorkbookStream = new FileInputStream(filePath)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedWorkbookStream, actualWorkbookStream)
  }

  "Exscalabur" should "write data in parts" in {
    val filePath = ExcelTestHelpers.createTempFile("sepRepRows").getAbsolutePath
    val animals = List("bear", "eagle", "elephant", "bird", "snake", "pig", "dog", "cat", "penguin", "anteater");
    val weight = List(1, 2, 3, 4, 5, 6, 7, 8, 9, 10);

    val repeatedData1 = animals.take(5)
      .map(animal => DataRow().addCell("animal", animal))

    val repeatedData2 = animals.slice(5, 10)
      .map(animal => DataRow().addCell("animal", animal))

    val repeatedData3 = weight.map(weight => DataRow().addCell("weight", weight))

    val schema = Map(
      "$REP.animal" -> YamlEntry(DataType.string, ExcelType.string),
      "$REP.weight" -> YamlEntry(DataType.long, ExcelType.number),
    )

    val templatePath = getClass.getResource("/excel/templates/seqRepRows.xlsx").getFile

    val exscalabur = Exscalabur(List(templatePath), schema, 100)

    val sheetWriter = exscalabur.getAppendOnlySheetWriter("Sheet1")

    sheetWriter.writeRepeatedData(repeatedData1)
    sheetWriter.writeRepeatedData(repeatedData2)
    sheetWriter.writeRepeatedData(repeatedData3)

    exscalabur.exportToFile(filePath)

    val expectedWorkbookStream = getClass.getResourceAsStream("/excel/expected/seqRepRows.xlsx")
    val actualWorkbookStream = new FileInputStream(filePath)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedWorkbookStream, actualWorkbookStream)
  }

  "Exscalabur" should "copy formula cells, shifting if needed" in {
    val filePath = ExcelTestHelpers.createTempFile("sepRepRows").getAbsolutePath

    val repeatedData1 = List("a", "b", "c").map(value => DataRow().addCell("repeated", value))
    val staticData = List(
      DataCell("static1", 1),
      DataCell("static2", 2),
      DataCell("static4", 3),
      DataCell("static5", 4),
    )

    val schema = Map(
      "$REP.repeated" -> YamlEntry(DataType.string, ExcelType.string),
      "$KEY.static1" -> YamlEntry(DataType.long, ExcelType.number),
      "$KEY.static2" -> YamlEntry(DataType.long, ExcelType.number),
      "$KEY.static4" -> YamlEntry(DataType.long, ExcelType.number),
      "$KEY.static5" -> YamlEntry(DataType.long, ExcelType.number),
    )

    val templatePath = getClass.getResource("/excel/templates/formulaSpec.xlsx").getFile

    val exscalabur = Exscalabur(List(templatePath), schema, 100)
    val sheetWriter = exscalabur.getAppendOnlySheetWriter("Sheet1")

    sheetWriter.writeRepeatedData(repeatedData1)
    sheetWriter.writeStaticData(staticData)

    exscalabur.exportToFile(filePath)

    val expectedWorkbookStream = getClass.getResourceAsStream("/excel/expected/formulaSpec.xlsx")
    val actualWorkbookStream = new FileInputStream(filePath)
    ExcelTestHelpers.assertEqualsWorkbooks(expectedWorkbookStream, actualWorkbookStream)
  }
}