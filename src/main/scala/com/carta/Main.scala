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
package com.carta

import java.io.File

import com.carta.exscalabur.{DataCell, DataRow, Exscalabur}
import com.carta.yaml.YamlReader

object Main extends App {
  override def main(args: Array[String]): Unit = {
    val templates: List[String] = "/home/katie/Documents/Exscalabur Demo/template.xlsx" :: Nil

    val yamlReader = YamlReader()
    val yamlData = yamlReader.parse(new File("/home/katie/Documents/Exscalabur Demo/template.yml"))

    val sword = Exscalabur(
      templates,
      yamlData,
      100
    )

    val repeatedData: List[DataRow] = List(
      DataRow().addCell("animal", "monkey"),
      DataRow().addCell("animal", "horse"),
      DataRow().addCell("animal", "cow"),
    )

    val repeatedData2 = List(
      DataRow().addCell("animal", "dog"),
      DataRow().addCell("animal", "cat"),
    )

    val repeatedData3 = List(
      DataRow().addCell("element", "hydrogen"),
      DataRow().addCell("element", "helium"),
      DataRow().addCell("element", "lithium")
    )

    val staticData = List(DataCell("name", "katie"))


    val sheetWriter = sword.getAppendOnlySheetWriter("Sheet1")

    sheetWriter.writeRepeatedData(repeatedData)
    sheetWriter.writeRepeatedData(repeatedData2)
    sheetWriter.writeRepeatedData(repeatedData2)
    sheetWriter.writeStaticData(staticData)
    sheetWriter.writeRepeatedData(repeatedData3)


    sword.exportToFile("/home/katie/Documents/Exscalabur Demo/out/exscalabur.xlsx")
  }
}
