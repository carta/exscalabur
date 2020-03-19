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
    }
  }.toMap

  private def writeDataToRows(staticDataModelMap: ModelMap, repeatedDataModelMaps: Seq[ModelMap]): Int = {
    val insertRowsWithData = insertRows(staticDataModelMap, repeatedDataModelMaps) _
    templateSheet.getRowIndices.drop(templateIndex).foldLeft(outputRowIndex)(insertRowsWithData)
  }

  private def insertRows(staticData: ModelMap, repeatedData: Iterable[ModelMap])
                        (writeRowIndex: Int, templateRowIndex: Int): Int = {
    Option(templateSheet.getRow(templateRowIndex)).map { templateRow =>
      if (isRepeatedRow(templateRow)) {
        repeatedData.filterNot(modelMap => shouldSkipRow(templateRow, modelMap, staticData))
          .foldLeft(writeRowIndex) { (currWriteIndex, modelMap) =>
            createRow(templateRow, modelMap, currWriteIndex)
            templateIndex = templateRowIndex
            currWriteIndex + 1
          }
      }
      else if (shouldSkipRow(templateRow, staticData)) {
        writeRowIndex
      }
      else {
        createRow(templateRow, staticData, writeRowIndex)
        templateIndex = templateRowIndex
        writeRowIndex + 1
      }
    } getOrElse {
      outputSheet.createRow(writeRowIndex)
      templateIndex = templateRowIndex + 1
      writeRowIndex + 1
    }
  }

  private def initialWrite(): Unit = {
    copyPictures()
    templateSheet.getRowIndices
      .drop(templateIndex)
      .foreach { currTemplateIndex =>
        Option(templateSheet.getRow(currTemplateIndex))
          .filter(row => !isRepeatedRow(row) && !shouldSkipRow(row, Map.empty))
          .foreach { templateRow =>
            createRow(templateRow, Map.empty, outputRowIndex)
            outputRowIndex += 1
            templateIndex += 1
          }
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