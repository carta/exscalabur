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
