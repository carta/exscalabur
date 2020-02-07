package main

import java.io.{FileInputStream, FileOutputStream}
import java.util.UUID.randomUUID

import exporters.report.utils.ExcelWorkbook
import resource.{ManagedResource, managed}

import scala.util.{Failure, Success}


object Writer {

  val template1 = "/Users/daviddufour/code/hackathon/exscalabur/src/resources/templates/tab1_template.xlsx"
  val template2 = "/Users/daviddufour/code/hackathon/exscalabur/src/resources/templates/tab2_template.xlsx"

  def writeExcelFileToDisk(windowSize: Int): Unit = {
    // this should be parameter, just write to test/UUID.xlsx for now
    val tempFilePath = s"/tmp/$randomUUID.xlsx"
    val tabNameToExcelTemplateMap = Map("tab1_name" -> template1,
                                        "tab2_name" -> template2)
    val tabNameToStreamMap = tabNameToExcelTemplateMap.map { entry =>
      (entry._1, managed(new FileInputStream(entry._2)))
    }
    val workbook = new ExcelWorkbook(tabNameToStreamMap, windowSize)
    val fos = new FileOutputStream(tempFilePath)

    workbook.copyAndSubstitute("tab1_name", ExportModelUtils.getModelMap(12L))

    writeRepeatedTabFuture(workbook, "tab2_name")

    // Writes the final workbook to the FileOutputStream with the given pathname, and then closes both the workbook and FileOutputStream
    workbook.write(fos)
    workbook.close()
    fos.close()
  }

  private def writeRepeatedTabFuture(workbook: ExcelWorkbook,
                                     tabName: String): Int = {
    val templateIndex = workbook.copyAndSubstitute(tabName)
    templateIndex match {
      case Some(startingIndex) =>
        workbook.insertRows(tabName,
                            startingIndex,
                            startingIndex,
                            List("a", "b", "c").map(ExportModelUtils.getModelMap)) match {
          case Success(value: Int) => value
          case Failure(exception) => throw exception
        }
      case None => throw new Exception(s"No repeated row for tab $tabName")
    }
  }
}