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
      DataCell("string_field", "katie"),
      DataCell("long_field", 1234),
      DataCell("double_field", 1235.1),
    )
    val staticData2 = List(
      DataCell("string_field2", "katie2"),
      DataCell("long_field2", 21234),
      DataCell("double_field2", 21235.1),
    )
    val sheetWriter = sword.getAppendOnlySheetWriter("Sheet1")

    val dataProvider: Iterator[(List[DataCell], List[DataRow])] = Iterator(
      (staticData1, List.empty),
      (staticData2, List.empty),
    )

    sheetWriter.writeData(dataProvider)

    sword.writeToDisk()
  }
}
