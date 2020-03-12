package com.carta.exscalabur

import java.io.File

import com.carta.excel.{SheetData, TabType, Writer}
import com.carta.yaml.{YamlEntry, YamlReader}

import scala.collection.mutable

//TODO verify no keys in yaml match
class Exscalabur(outputPath: String, yamlPath: String) {
  private val tabToRows = mutable.Map.empty[String, List[DataRow]]
  private val tabToTemplatePath = mutable.Map.empty[String, String]
  private val templateToStuff = mutable.Map.empty[String, (String, Iterable[DataCell], List[DataRow])]

  private val yamlReader = YamlReader()
  private lazy val yamlData = yamlReader.parse(new File(yamlPath))

  val windowSize = 100

  val writer = new Writer(windowSize)

  // TODO: I think this is actually add template
  def addTab(tabName: String, templatePath: String, staticData: Iterable[DataCell], repeatedData: List[DataRow]): Exscalabur = {
    tabToRows.put(tabName, repeatedData)
    tabToTemplatePath.put(tabName, templatePath)
    templateToStuff.put(tabName, (templatePath, staticData, repeatedData))
    this
  }

  def writeExcelToDisk(): Unit = {
    writer.writeExcelFileToDisk(outputPath, getTabParams(TabType.repeated, yamlData))
  }

  private def getTabParams(tabType: TabType.Value, yamlData: Map[String, YamlEntry]) = {
    val toTabParam = ((tabName: String, data: Iterable[DataCell], repeatedData: List[DataRow], templatePath: String) => {
      SheetData(tabName, templatePath, data, repeatedData, yamlData)
    }).tupled

    templateToStuff.map { z =>
      toTabParam(z._1, z._2._2, z._2._3, z._2._1)
    }
  }
}
