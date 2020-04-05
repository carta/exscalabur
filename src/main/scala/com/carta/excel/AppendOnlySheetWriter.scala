package com.carta.excel

import java.util.Date

import com.carta.excel.ExportModelUtils.ModelMap
import com.carta.excel.implicits.ExtendedSheet._
import com.carta.exscalabur
import com.carta.exscalabur.{CellType => _, _}
import com.carta.yaml.{YamlCellType, YamlEntry}
import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.{XSSFPicture, XSSFSheet}

import scala.collection.JavaConverters._
import scala.collection.mutable

class AppendOnlySheetWriter
(templateSheet: XSSFSheet,
 outputWorkbook: Workbook,
 schema: Map[String, YamlEntry],
 cellStyleCache: mutable.Map[CellStyle, Int]) {
  var outputRowIndex: Int = templateSheet.getFirstRowNum
  var templateIndex: Int = templateSheet.getFirstRowNum
  val outputSheet: Sheet = outputWorkbook.createSheet(templateSheet.getSheetName)
  val test = mutable.Map.empty[Int, Boolean]
  def copyPictures(): Unit = {
    val templateImages = templateSheet.createDrawingPatriarch().getShapes.asScala
      .filter(_.isInstanceOf[XSSFPicture])
      .map(_.asInstanceOf[XSSFPicture])

    templateImages.foreach { img: XSSFPicture =>
      val drawing = outputSheet.createDrawingPatriarch()
      val anchor = outputSheet.getWorkbook.getCreationHelper.createClientAnchor()
      anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE)
      val imgData = img.getPictureData
      val templateAnchor = img.getClientAnchor
      val picIndex = outputWorkbook.addPicture(imgData.getData, imgData.getPictureType)

      anchor.setCol1(templateAnchor.getCol1)
      anchor.setCol2(templateAnchor.getCol2)
      anchor.setRow1(templateAnchor.getRow1)
      anchor.setRow2(templateAnchor.getRow2)
      anchor.setDx1(templateAnchor.getDx1)
      anchor.setDx2(templateAnchor.getDx2)
      anchor.setDy1(templateAnchor.getDy1)
      anchor.setDy2(templateAnchor.getDy2)

      drawing.createPicture(anchor, picIndex)
    }
  }

  def writeData(dataProvider: Iterator[(Iterable[DataCell], Seq[DataRow])]): Unit = {
    dataProvider.foreach { case (staticData, repeatedData) =>
      val staticDataModelMap = getModelMap(ExportModelUtils.SUBSTITUTION_KEY, staticData)
      val repeatedDataModelMaps = repeatedData.map { dataRow =>
        getModelMap(ExportModelUtils.REPEATED_FIELD_KEY, dataRow.getCells)
      }
      outputRowIndex = writeDataToRows(staticDataModelMap, repeatedDataModelMaps)
    }
  }

  private def getModelMap(keyType: String, dataRow: Iterable[DataCell]): ModelMap = {
    dataRow.map(_.asTuple).map { case (key: String, value: exscalabur.CellType) =>
      val newKey = s"${keyType}.$key"
      val newValue = schema(newKey) match {
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
    }.toMap
  }

  private def writeDataToRows(staticData: ModelMap, repeatedData: Seq[ModelMap]): Int = {
    templateSheet.getRowIndices
      .drop(templateIndex) // Append-Only Sheet Writer does not support writing data to previously written template rows
      .foldLeft(outputRowIndex) { (writeRowIndex, templateRowIndex) =>
        templateSheet.rowOpt(templateRowIndex)
          .map { templateRow =>
            val writeResult: RowWriteResult = writeDataToRow(templateRow, writeRowIndex, repeatedData, staticData)
            val nextWriteIndex = writeResult.writeIndex
            if (writeResult.shouldStopCurrentWrite) {
              return nextWriteIndex
            }
            else {
              nextWriteIndex
            }
          }
          .getOrElse {
            outputSheet.createRow(writeRowIndex)
            templateIndex = templateRowIndex + 1
            writeRowIndex + 1
          }
      }
  }

  case class RowWriteResult(writeIndex: Int, shouldStopCurrentWrite: Boolean)

  private def writeDataToRow(templateRow: Row, writeRowIndex: Int, repeatedData: Seq[ModelMap], staticData: ModelMap): RowWriteResult = {
    if (isRepeatedRow(templateRow)) {
      val x = writeRepeatedDataToRow(templateRow, writeRowIndex, repeatedData, staticData)
      println(x)
      return x
    }
    else if (shouldSkipRow(templateRow, staticData)) {
      println("SKIPPING ROW")
      RowWriteResult(writeRowIndex, shouldStopCurrentWrite = true)
    }
    else {
      createRow(templateRow, staticData, writeRowIndex)
      templateIndex = templateRow.getRowNum + 1
      RowWriteResult(writeRowIndex + 1, shouldStopCurrentWrite = false)
    }
  }

  private def writeRepeatedDataToRow(templateRow: Row, writeRowIndex: Int, repeatedData: Seq[ModelMap], staticData: ModelMap): RowWriteResult = {
    val dataForRow = repeatedData.zipWithIndex
      .filterNot { case (modelMap, _) => shouldSkipRow(templateRow, modelMap, staticData) }
      .map { case (modelMap, rowIndex) => (modelMap, rowIndex == repeatedData.size - 1) }
    if (dataForRow.isEmpty) {
      templateIndex = templateRow.getRowNum + (if (test.getOrElse(templateRow.getRowNum, false)) 1 else 0)
      return RowWriteResult(writeRowIndex, shouldStopCurrentWrite = !test.getOrElse(templateRow.getRowNum, false))
    }
    val writeIndex = dataForRow.foldLeft(writeRowIndex) { case (currWriteIndex, (modelMap, isLastDataMap)) =>
      createRow(templateRow, modelMap, currWriteIndex)
      templateIndex = templateRow.getRowNum
      test.put(templateIndex, true)
      lazy val shouldSkipNextRow = templateSheet.rowOpt(templateRow.getRowNum + 1).forall(row => shouldSkipRow(row, staticData))
      if (isLastDataMap && shouldSkipNextRow) {
        return RowWriteResult(currWriteIndex + 1, shouldStopCurrentWrite = true)
      }
      else {
        currWriteIndex + 1
      }
    }
    RowWriteResult(writeIndex, shouldStopCurrentWrite = dataForRow.isEmpty)
  }

  private def initialWrite(): Unit = {
    copyPictures()
    templateSheet.rowIterator().asScala
      .find(row => !isRepeatedRow(row) && !shouldSkipRow(row, Map.empty))
      .foreach { templateRow =>
        createRow(templateRow, Map.empty, outputRowIndex)
        outputRowIndex += 1
        templateIndex += 1
      }
  }

  private def isRepeatedRow(row: Row): Boolean = row.cellIterator().asScala.exists(isRepeatedCell)

  private def isSubstituteCell(cell: Cell): Boolean = {
    cell.getCellType == CellType.STRING &&
      (cell.getStringCellValue.startsWith(ExportModelUtils.REPEATED_FIELD_KEY) ||
        cell.getStringCellValue.startsWith(ExportModelUtils.SUBSTITUTION_KEY))
  }

  private def isRepeatedCell(cell: Cell): Boolean = {
    cell.getCellType == CellType.STRING &&
      cell.getStringCellValue.startsWith(ExportModelUtils.REPEATED_FIELD_KEY)
  }

  private def createRow(templateRow: Row, modelMap: ModelMap, writeIndex: Int): Unit = {
    val outputRow = outputSheet.createRow(writeIndex)
    populateRowFromTemplate(templateRow, outputRow, templateSheet, outputSheet, modelMap)
  }

  private def populateRowFromTemplate(templateRow: Row,
                                      outputRow: Row,
                                      templateSheet: Sheet,
                                      outputSheet: Sheet,
                                      substitutionMap: ExportModelUtils.ModelMap
                                     ): Unit = {

    copyRowHeight(templateRow, outputRow)

    templateRow.cellIterator().asScala
      .foreach { cell =>
        val cellIndex = cell.getColumnIndex
        val outputCell = outputRow.createCell(cellIndex)
        copyColumnWidth(templateSheet, outputSheet, cellIndex)
        applyCellStyleFromTemplate(cell, outputCell)
        substituteAndCopyCell(cell, outputCell, substitutionMap)
      }
  }

  private def copyRowHeight(from: Row, to: Row): Unit = {
    to.setHeight(from.getHeight)
  }

  private def copyColumnWidth(from: Sheet, to: Sheet, col: Int): Unit = {
    to.setColumnWidth(col, from.getColumnWidth(col))
  }

  private def applyCellStyleFromTemplate(from: Cell, to: Cell): Unit = {
    val templateStyle = from.getCellStyle

    def putCellStyle = {
      val outputStyle = outputWorkbook.createCellStyle()
      outputStyle.cloneStyleFrom(templateStyle)
      outputStyle.getIndex
    }

    to.setCellStyle(outputWorkbook.getCellStyleAt(cellStyleCache.getOrElseUpdate(templateStyle, putCellStyle)))
  }

  private def substituteAndCopyCell(templateCell: Cell, outputCell: Cell, substitutionMap: ExportModelUtils.ModelMap): Unit = {
    templateCell.getCellType match {
      case CellType.NUMERIC =>
        outputCell.setCellValue(templateCell.getNumericCellValue)
      case CellType.BOOLEAN =>
        outputCell.setCellValue(templateCell.getBooleanCellValue)
      case CellType.STRING if isSubstituteCell(templateCell) =>
        substituteString(templateCell, outputCell, substitutionMap)
      case CellType.STRING =>
        outputCell.setCellValue(templateCell.getStringCellValue)
      case _ =>
      //logger.trace("CellType is one of {CellType.ERROR, CellType.BLANK, CellType.FORMULA, CellType.NONE} which we do not want to copy from")
    }
  }

  private def substituteString(templateCell: Cell, outputCell: Cell, substitutionMap: ExportModelUtils.ModelMap): Unit = {
    val templateValue = templateCell.getStringCellValue
    substitutionMap.get(templateValue).foreach {
      case CellDouble(double: Double) => outputCell.setCellValue(double)
      case CellDate(date: Date) => outputCell.setCellValue(date)
      case CellString(string: String) => outputCell.setCellValue(string)
      case _ => throw new NoSuchElementException(s"Unsupported cell type for key: $templateValue")
    }
  }

  private def shouldSkipRow(templateRow: Row, modelMap: ModelMap*): Boolean = {
    !templateRow.cellIterator().asScala
      .filter(isSubstituteCell)
      .map(_.getStringCellValue)
      .forall(cellValue => modelMap.exists(modelMap => modelMap.contains(cellValue)))
  }
}

object AppendOnlySheetWriter {
  def apply(templateSheet: XSSFSheet, outputWorkbook: SXSSFWorkbook, schema: Map[String, YamlEntry], cellStyleCache: mutable.Map[CellStyle, Int]): AppendOnlySheetWriter = {
    val sheetWriter = new AppendOnlySheetWriter(templateSheet, outputWorkbook, schema, cellStyleCache)
    sheetWriter.initialWrite()
    sheetWriter
  }
}