package com.carta.excel

import java.io.{FileInputStream, FileOutputStream, OutputStream}

import com.carta.excel.ExportModelUtils.ModelMap
import com.carta.exscalabur._
import com.carta.yaml.{CellType, YamlEntry}
import resource.{ManagedResource, managed}


case class SheetData(sheetName: String,
                     templatePath: String,
                     staticData: DataRow,
                     repeatedData: List[DataRow],
                     schema: Map[String, YamlEntry]) {
  def toStreamTuple: (String, ManagedResource[FileInputStream]) = {
    (sheetName, managed(new FileInputStream(templatePath)))
  }
}

class Writer(val windowSize: Int) {
  def writeExcelToStream(sheets: Iterable[SheetData], outputStream: OutputStream): Unit = {
    val sheetNameToStreamMap = sheets.map(_.toStreamTuple).toMap
    val workbook = new ExcelWorkbook(sheetNameToStreamMap, windowSize)

    sheets.foreach {
      case SheetData(templateName, _, tabData, repeatedTabData, tabSchema) =>
        val staticDataModelMap = getModelMap(ExportModelUtils.SUBSTITUTION_KEY, tabSchema, tabData)
        val repeatedDataModelMap = repeatedTabData.map(rowData => getModelMap(ExportModelUtils.REPEATED_FIELD_KEY, tabSchema, rowData))

        val (sheetName, startIndexOpt) = workbook.copyAndSubstitute(templateName, staticDataModelMap).head

        startIndexOpt.map(index => workbook.insertRows(templateName, index, sheetName, index, repeatedDataModelMap))
    }
    //TODO this returns a Try[Unit] -- error handling
    workbook.write(outputStream)
    workbook.close()
  }

  def writeExcelFileToDisk(filePath: String, sheets: Iterable[SheetData]): Unit = {
    val outputStream = new FileOutputStream(filePath)
    writeExcelToStream(sheets, outputStream)
    outputStream.close()
  }

  private def getModelMap(keyType: String, tabSchema: Map[String, YamlEntry], dataRow: DataRow): ModelMap = {
    dataRow.data.map { case (key: String, value: CellType) =>
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
}