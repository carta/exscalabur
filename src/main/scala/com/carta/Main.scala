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
      "/home/katie/Documents/Exscalabur Demo/out/exscalabur.xlsx",
      yamlData,
      100
    )

    val repeatedData: List[DataRow] = List(
      DataRow.Builder().addCell("animal", "monkey").build(),
      DataRow.Builder().addCell("animal", "horse").build(),
      DataRow.Builder().addCell("animal", "cow").build(),
    )

    val repeatedData2 = List(
      DataRow.Builder().addCell("animal", "dog").build(),
      DataRow.Builder().addCell("animal", "cat").build(),
    )

    val repeatedData3 = List(
      DataRow.Builder().addCell("element", "hydrogen").build(),
      DataRow.Builder().addCell("element", "helium").build(),
      DataRow.Builder().addCell("element", "lithium").build()
    )


    val sheetWriter = sword.getAppendOnlySheetWriter("Sheet1")
    val dataProvider: Iterator[(List[DataCell], List[DataRow])] = Iterator(
      (List.empty, repeatedData),
      (List.empty, repeatedData2),
      (List.empty, repeatedData2),
      (List(DataCell("name", "katie")), List.empty),
      (List.empty, repeatedData3)
    )


    sheetWriter.writeData(dataProvider)

    sword.writeToDisk()
  }
}
