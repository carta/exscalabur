package com.carta.excel

import java.io.{FileInputStream, FileOutputStream}
import java.util.UUID.randomUUID

import com.carta.excel.ExportModelUtils.ModelMap
import com.carta.excel._
import resource.{ManagedResource, managed}

import scala.util.{Failure, Success}


object Writer {

  val template1 = "/Users/daviddufour/code/hackathon/exscalabur/src/resources/templates/tab1_template.xlsx"
  val template2 = "/Users/daviddufour/code/hackathon/exscalabur/src/resources/templates/tab2_template.xlsx"

  def writeExcelFileToDisk(windowSize: Int): Unit = {
    // this should be parameter, just write to test/UUID.xlsx for now
    val tempFilePath = s"/tmp/$randomUUID.xlsx"
    val tabNameToExcelTemplateMap = Map("template1" -> template1,
                                        "template2" -> template2)
    val tabNameToStreamMap = tabNameToExcelTemplateMap.map { entry =>
      (entry._1, managed(new FileInputStream(entry._2)))
    }
    val workbook = new ExcelWorkbook(tabNameToStreamMap, windowSize)
    val fos = new FileOutputStream(tempFilePath)

    workbook.copyAndSubstitute("template1", ExportModelUtils.getModelMap(12L))

    val repeatedRowIndices = workbook.copyAndSubstitute("template2", ExportModelUtils.getModelMap(24L))
    writeRepeatedTabFuture(
      tabName = repeatedRowIndices(0)._1,
      repeatedRowIndex = repeatedRowIndices(0)._2,
      workbook = workbook,
      templateName = "template2",
      substitutionMaps = List("a", "b", "c").map(ExportModelUtils.getModelMap))

    // Writes the final workbook to the FileOutputStream with the given pathname, and then closes both the workbook and FileOutputStream
    workbook.write(fos)
    workbook.close()
    fos.close()
  }

  private def writeRepeatedTabFuture(tabName: String,
                                     repeatedRowIndex: Option[Int],
                                     workbook: ExcelWorkbook,
                                     templateName: String,
                                     substitutionMaps: List[ModelMap]): Int = {
    repeatedRowIndex match {
      case Some(startingIndex) =>
        // TODO: Allow users to call this again with startingIndex and the return value from insertRows
        workbook.insertRows(templateName,
                            startingIndex,
                            tabName,
                            startingIndex,
                            substitutionMaps) match {
          case Success(value: Int) => value
          case Failure(exception) => throw exception
        }
      case None => throw new Exception(s"No repeated row for template: $templateName")
    }
  }
}