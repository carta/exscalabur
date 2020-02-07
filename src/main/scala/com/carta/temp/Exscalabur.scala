package com.carta.temp

import java.io.File

import com.carta.excel.{TabParam, TabType, Writer}
import com.carta.yaml.{KeyObject, YamlReader}

import scala.collection.mutable

//TODO verify no keys in yaml match
class Exscalabur(outputPath: String, yamlPath: String) {
  private val tabToRows = mutable.Map.empty[String, List[DataRow]]
  private val tabToTemplatePath = mutable.Map.empty[String, String]

  private val yamlReader = new YamlReader()
  private lazy val yamlData = yamlReader.parse(new File(yamlPath))

  val windowSize = 100;

  def addTab(tabName: String, templatePath: String, data: List[DataRow]): Exscalabur = {
    tabToRows.put(tabName, data)
    tabToTemplatePath.put(tabName, templatePath)
    this
  }

  def writeExcelToDisk(): Unit = {
    Writer.writeExcelFileToDisk(outputPath, windowSize, getTabParams(TabType.repeated, yamlData))
  }

  private def getTabParams(tabType: TabType.Value, yamlData: Map[String, KeyObject]) = {
    val toTabParam = ((tabName: String, data: List[DataRow], templatePath: String) => {
      TabParam(tabType, tabName, templatePath, data, yamlData)
    }).tupled

    (tabToRows.keySet, tabToRows.values, tabToTemplatePath.values)
      .zipped
      .toList
      .map(toTabParam)
  }
}
