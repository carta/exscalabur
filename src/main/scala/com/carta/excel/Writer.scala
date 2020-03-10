package com.carta.excel

import java.io.{FileInputStream, FileOutputStream}

import com.carta.excel.ExportModelUtils.ModelMap
import com.carta.exscalabur._
import com.carta.yaml.{CellType, YamlEntry}
import resource.{ManagedResource, managed}


case class TabParam(tabName: String,
                    templatePath: String,
                    tabData: DataRow,
                    repeatedTabData: List[DataRow],
                    tabSchema: Map[String, YamlEntry]) {
  def toStreamTuple: (String, ManagedResource[FileInputStream]) = {
    (tabName, managed(new FileInputStream(templatePath)))
  }
}

object Writer {

  val template1 = "/Users/daviddufour/code/hackathon/exscalabur/src/resources/templates/tab1_template.xlsx"
  val template2 = "/Users/daviddufour/code/hackathon/exscalabur/src/resources/templates/tab2_template.xlsx"

  //TODO: restrict to supported types
  def writeExcelFileToDisk(filePath: String, windowSize: Int, tabs: Iterable[TabParam]): Unit = {
    // this should be parameter, just write to test/UUID.xlsx for now
    val tabNameToStreamMap = tabs.map(_.toStreamTuple).toMap

    val workbook = new ExcelWorkbook(tabNameToStreamMap, windowSize)
    val fos = new FileOutputStream(filePath)

    //    workbook.copyAndSubstitute("tab1_name", ExportModelUtils.getModelMap(12L))
    //
    //    val repeatedRowIndex = workbook.copyAndSubstitute("tab2_name", ExportModelUtils.getModelMap(24L))
    //    writeRepeatedTabFuture(
    //      repeatedRowIndex = repeatedRowIndex,
    //      workbook = workbook,
    //      tabName = "tabe2_name",
    //      maps = List("a", "b", "c").map(ExportModelUtils.getModelMap))


    tabs.foreach {
      case TabParam(templateName: String,
                    _,
                    tabData: DataRow,
                    repeatedTabData: List[DataRow],
                    tabSchema: Map[String, YamlEntry]) =>
        //TODO proper error handling on copyAndSubstitute
        val startIndex = workbook.copyAndSubstitute(templateName, getModelMap(ExportModelUtils.SUBSTITUTION_KEY, tabSchema, tabData)).head
        val modelMaps = repeatedTabData.map(rowData => getModelMap(ExportModelUtils.REPEATED_FIELD_KEY, tabSchema, rowData))
        startIndex._2 match {
          case Some(index) => workbook.insertRows(templateName, index, startIndex._1, index, modelMaps)
          case None =>
        }
    }
    // Writes the final workbook to the FileOutputStream with the given pathname, and then closes both the workbook and FileOutputStream
    workbook.write(fos)
    workbook.close()
    fos.close()
  }

  private def getModelMap(keyType: String, tabSchema: Map[String, YamlEntry], dataRow: DataRow): ModelMap = {
    dataRow.data.map { case (key: String, value: CellType) =>
      println(tabSchema)
      val newKey = s"${keyType}.$key"
      val newValue = tabSchema(newKey) match {
        // TODO different input output types
        case YamlEntry(_, CellType.string, CellType.string) =>
          ExportModelUtils.toCellStringFromString(value.asInstanceOf[StringCellType].value)
        case YamlEntry(_, CellType.double, CellType.double) =>
          ExportModelUtils.toCellDoubleFromDouble(value.asInstanceOf[DoubleCellType].value)
        case YamlEntry(_, CellType.long, CellType.double) =>
          ExportModelUtils.toCellDoubleFromLong(value.asInstanceOf[LongCellType].value)
        case YamlEntry(_, CellType.long, CellType.date) =>
          ExportModelUtils.toCellDateFromLong(value.asInstanceOf[LongCellType].value)
      }
      newKey -> newValue
    }
  }

  //  private def writeRepeatedTabFuture(repeatedRowIndex: Option[Int], workbook: ExcelWorkbook, tabName: String, maps: List[ModelMap]): Int = {
  //    repeatedRowIndex match {
  //      case Some(startingIndex) =>
  //        workbook.insertRows(tabName,
  //                            startingIndex,
  //                            startingIndex,
  //                            maps) match {
  //          case Success(value: Int) => value
  //          case Failure(exception) => throw exception
  //        }
  //      case None => throw new Exception(s"No repeated row for tab $tabName")
  //    }
  //  }
}