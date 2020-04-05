package com.carta

import java.io.File

import com.carta.exscalabur.{DataCell, DataRow, Exscalabur}
import com.carta.yaml.YamlReader

object Main extends App {
  override def main(args: Array[String]): Unit = {
    val templates: List[String] = "/Users/katiedsouza/Desktop/demo.xlsx" :: Nil

    val yamlReader = YamlReader()
    val yamlData = yamlReader.parse(new File("/Users/katiedsouza/Desktop/demo.yaml"))

    val sword = Exscalabur(
      templates,
      "/Users/katiedsouza/Desktop/demo.out.xlsx",
      yamlData,
      100
    )

    val staticData1 = List(
      DataCell("fname", "katie"),
      DataCell("lname", "dsouza"),
    )

    val repeatedData1 = List(
      DataRow.Builder().addCell("animal", "monkey").addCell("weight", 12.1).build(),
      DataRow.Builder().addCell("animal", "horse").addCell("weight", 12.2).build()
    )

    val staticData2 = List(
      DataCell("conclusion", "EXSCALABUR")
    )

    val repeatedData2 = List(
      DataRow.Builder().addCell("element", "hydrogen").build(),
      DataRow.Builder().addCell("element", "helium").build(),
      DataRow.Builder().addCell("element", "lithium").build()
    )


    val sheetWriter = sword.getAppendOnlySheetWriter("Sheet1")

    val dataProvider: Iterator[(List[DataCell], List[DataRow])] = Iterator(
      (staticData1, List.empty),
      (List.empty, repeatedData1),
      (staticData2, List.empty),
      (List.empty, repeatedData2)
    )


    sheetWriter.writeData(dataProvider)

    sword.writeToDisk()
  }
}
