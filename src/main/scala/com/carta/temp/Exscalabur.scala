package com.carta.temp

import java.io.File

import com.carta.excel.{TabParam, TabType, Writer}
import com.carta.yaml.{YamlEntry, YamlReader}

import scala.collection.{immutable, mutable}

//TODO verify no keys in yaml match
class Exscalabur(outputPath: String, yamlPath: String) {
  private val tabToRows = mutable.Map.empty[String, List[DataRow]]
  private val tabToTemplatePath = mutable.Map.empty[String, String]
  private val templateToStuff = mutable.Map.empty[String, (String, DataRow, List[DataRow])]

  private val yamlReader = YamlReader()
  private lazy val yamlData = yamlReader.parse(new File(yamlPath))

  val windowSize = 100

  // TODO: I think this is actually add template
  def addTab(tabName: String, templatePath: String, data: DataRow, repeatedData: List[DataRow]): Exscalabur = {
    tabToRows.put(tabName, repeatedData)
    tabToTemplatePath.put(tabName, templatePath)
    templateToStuff.put(tabName, (templatePath, data, repeatedData))
    this
  }

  def writeExcelToDisk(): Unit = {
    Writer.writeExcelFileToDisk(outputPath, windowSize, getTabParams(TabType.repeated, yamlData))
  }

  private def getTabParams(tabType: TabType.Value, yamlData: Map[String, YamlEntry]) = {
    val toTabParam = ((tabName: String, data: DataRow, repeatedData: List[DataRow], templatePath: String) => {
      TabParam(tabName, templatePath, data, repeatedData, yamlData)
    }).tupled

    templateToStuff.map { z =>
      toTabParam(z._1, z._2._2, z._2._3, z._2._1)
    }
  }
}
