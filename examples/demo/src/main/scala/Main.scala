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
import java.time.LocalDate

import com.carta.exscalabur.{DataCell, DataRow, Exscalabur}
import com.carta.yaml.YamlReader

object Main {
  def main(args: Array[String]): Unit = {
    val templates: List[String] = List(getClass.getResource("/excel/template.xlsx").getPath)

    val schemaPath = getClass.getResource("/yaml/schema.yaml").getPath
    val schemaDefinition = YamlReader().parse(schemaPath)

    val sword = Exscalabur(
      templates,
      schemaDefinition,
      windowSize = 100
    )

    val peopleData1: List[DataRow] = List(
      DataRow().addCell("person", "Jon Somebody").addCell("gpa", 3.8).addCell("major", "CS"),
      DataRow().addCell("person", "Jane Person").addCell("gpa", 4.2).addCell("major", "Engineering"),
      DataRow().addCell("person", "Frank LastName").addCell("gpa", 4.1).addCell("major", "Arts"),
    )

    val peopleData2: List[DataRow] = List(
      DataRow().addCell("person", "Sarah Fakename").addCell("gpa", 3.9).addCell("major", "Science"),
      DataRow().addCell("person", "Tim Tom").addCell("gpa", 3.9).addCell("major", "Engineering"),
    )

    val schoolData: List[DataRow] = List(
      DataRow().addCell("school", "School Name").addCell("numStudents", 43000),
      DataRow().addCell("school", "University of Place").addCell("numStudents", 25000),
      DataRow().addCell("school", "Well Known School").addCell("numStudents", 30000)
    )

    val recordHeaders: List[DataRecord] = List(
      DataRecord()
        .addCell("dateHeader", "Date")
        .addRecord(List(
          DataCell("peerCloseHeader", "AAPL close"),
          DataCell("peerChangeHeader", "AAPL")
          ))
        .addRecord(List(
          DataCell("peerCloseHeader", "MSFT close"),
          DataCell("peerChangeHeader", "MSFT")
          ))
        .addRecord(List(
          DataCell("peerCloseHeader", "GOOG close"),
          DataCell("peerChangeHeader", "GOOG")
          ))

    val recordData: List[DataRecord] = List(
      DataRecord()
        .addCell("date", LocalDate.of(2020, MAY, 1))
        .addRecord(List(
          DataCell("peerClose", 1.08),
          DataCell("peerChange", 0.01)
          ))
        .addRecord(List(
          DataCell("peerClose", 2.31),
          DataCell("peerChange", -0.01)
          ))
        .addRecord(List(
          DataCell("peerClose", 4.25),
          DataCell("peerChange", 0.02)
          )),
      DataRecord()
        .addCell("date", LocalDate.of(2020, JUNE, 1))
        .addRecord(List(
          DataCell("peerClose", 1.18),
          DataCell("peerChange", 0.03)
          ))
        .addRecord(List(
          DataCell("peerClose", 2.41),
          DataCell("peerChange", 0.13)
          ))
        .addRecord(List(
          DataCell("peerClose", 4.55),
          DataCell("peerChange", -0.03)
          ))
      )

    val staticData1 = List(DataCell("project", "exscalabur demo"))

    val staticData2 = List(
      DataCell("value1", 12.34),
      DataCell("value2", 11.35),
      DataCell("value3", 14.50),
    )
    val sheetWriter = sword.getAppendOnlySheetWriter("Sheet1")

    sheetWriter.writeStaticData(staticData1)
    sheetWriter.writeRepeatedData(peopleData1)
    sheetWriter.writeRepeatedData(peopleData2)
    sheetWriter.writeRepeatedData(schoolData)
    sheetWriter.writeRecordData(recordData)
    sheetWriter.writeStaticData(staticData2)

    sword.exportToFile("./out.xlsx")
  }
}