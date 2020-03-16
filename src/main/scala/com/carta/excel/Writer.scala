package com.carta.excel

import java.io.{FileInputStream, FileOutputStream, OutputStream}

import com.carta.excel.ExportModelUtils.ModelMap
import com.carta.excel.implicits.ExtendedSheet._
import com.carta.exscalabur._

import scala.collection.JavaConverters._
import com.carta.yaml.{YamlCellType, YamlEntry}
import resource.{ManagedResource, managed}

import scala.util.{Failure, Success}


case class SheetData(sheetName: String,
                     templatePath: String,
                     staticData: Iterable[DataCell],
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
        val repeatedDataModelMap = repeatedTabData.map { rowData =>
          getModelMap(ExportModelUtils.REPEATED_FIELD_KEY, tabSchema, rowData) ++ staticDataModelMap
        }

        workbook.sheets(templateName).foreach { templateSheet =>
          var outputRow = templateSheet.getFirstRowNum
          templateSheet.getRowIndices
            .foreach { rowIndex =>
              workbook.insertRows(templateName, rowIndex, templateSheet.getSheetName, outputRow, staticDataModelMap, repeatedDataModelMap) match {
                case Success(nextRow) =>  outputRow = nextRow
                case Failure(e) => //log this
              }
            }
        }
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
    this.getModelMap(keyType, tabSchema, dataRow.getCells)
  }

  private def getModelMap(keyType: String, tabSchema: Map[String, YamlEntry], dataRow: Iterable[DataCell]): ModelMap = {
    dataRow.map(_.asTuple).map { case (key: String, value: CellType) =>
      val newKey = s"${keyType}.$key"
      val newValue = tabSchema(newKey) match {
        // TODO different input output types
        case YamlEntry(_, YamlCellType.string, YamlCellType.string) =>
          ExportModelUtils.toCellStringFromString(value.asInstanceOf[StringCellType].value)
        case YamlEntry(_, YamlCellType.double, YamlCellType.double) =>
          ExportModelUtils.toCellDoubleFromDouble(value.asInstanceOf[DoubleCellType].value)
        case YamlEntry(_, YamlCellType.long, YamlCellType.double) =>
          ExportModelUtils.toCellDoubleFromLong(value.asInstanceOf[LongCellType].value)
        case YamlEntry(_, YamlCellType.long, YamlCellType.date) =>
          ExportModelUtils.toCellDateFromLong(value.asInstanceOf[LongCellType].value)
      }
      newKey -> newValue
    }
  }.toMap
}