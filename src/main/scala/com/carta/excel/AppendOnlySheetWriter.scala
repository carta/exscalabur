/**
 * Copyright 2018 eShares, Inc. dba Carta, Inc.
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * https://www.apache.org/licenses/LICENSE-2.0
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package com.carta.excel

import java.util.Date

import com.carta.excel.ExportModelUtils.ModelMap
import com.carta.excel.implicits.ExtendedRow._
import com.carta.excel.implicits.ExtendedSheet._
import com.carta.exscalabur._
import com.carta.yaml.YamlEntry
import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFSheet

import scala.collection.JavaConverters._
import scala.collection.mutable

class AppendOnlySheetWriter
(templateSheet: XSSFSheet,
 outputWorkbook: Workbook,
 schema: Map[String, YamlEntry],
 cellStyleCache: mutable.Map[CellStyle, Int]) {
  private var numRowsCreated = 0
  private var outputRowIndex: Int = templateSheet.getFirstRowNum
  private var templateIndex: Int = templateSheet.getFirstRowNum
  private val outputSheet: Sheet = outputWorkbook.createSheet(templateSheet.getSheetName)

  private val repeatedRowsWithPartialWrittenData = mutable.Set.empty[Int]


  private val cellFormulaParser = new CellFormulaParser()
  def copyPictures(): Unit = {
    templateSheet.getPictures.foreach(outputSheet.copyPicture)
  }

  def writeStaticData(data: Seq[DataCell]): Unit = {
    val staticDataMap = ModelMap(ExportModelUtils.SUBSTITUTION_KEY, data, schema)
    outputRowIndex = writeDataToRows(staticData = staticDataMap)
  }

  def writeRepeatedData(data: Seq[DataRow]): Unit = {
    val repeatedDataModelMaps: Seq[ModelMap] = data.map { dataRow =>
      ModelMap(ExportModelUtils.REPEATED_FIELD_KEY, dataRow.getCells, schema)
    }
    outputRowIndex = writeDataToRows(repeatedData = repeatedDataModelMaps)
  }

  private def writeDataToRows(staticData: ModelMap = Map.empty, repeatedData: Seq[ModelMap] = Seq.empty): Int = {
    templateSheet.getRowIndices
      .drop(templateIndex)
      .foldLeft(outputRowIndex) { (writeRowIndex, templateRowIndex) =>
        templateSheet.rowOpt(templateRowIndex)
          .map { templateRow =>
            val writeResult: RowWriteResult = writeDataToRow(
              templateRow = templateRow,
              writeRowIndex = writeRowIndex,
              staticData = staticData,
              repeatedData = repeatedData
            )
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
            numRowsCreated += 1
            templateIndex = templateRowIndex + 1
            writeRowIndex + 1
          }
      }
  }

  private case class RowWriteResult(writeIndex: Int, shouldStopCurrentWrite: Boolean)

  private def writeDataToRow(templateRow: Row,
                             writeRowIndex: Int,
                             repeatedData: Seq[ModelMap] = Seq.empty,
                             staticData: ModelMap = Map.empty): RowWriteResult = {
    if (templateRow.isRepeatedRow) {
      writeRepeatedDataToRow(templateRow, writeRowIndex, repeatedData, staticData)
    }
    else if (shouldSkipRow(templateRow, staticData)) {
      RowWriteResult(writeRowIndex, shouldStopCurrentWrite = true)
    }
    else {
      createRow(templateRow, staticData, writeRowIndex)
      templateIndex = templateRow.getRowNum + 1
      RowWriteResult(writeRowIndex + 1, shouldStopCurrentWrite = false)
    }
  }

  private def writeRepeatedDataToRow(templateRow: Row, writeRowIndex: Int, repeatedData: Seq[ModelMap], staticData: ModelMap): RowWriteResult = {
    templateIndex = templateRow.getRowNum
    val dataForRow = repeatedData.zipWithIndex
      .filterNot { case (modelMap, _) => shouldSkipRow(templateRow, modelMap, staticData) }
      .map { case (modelMap, rowIndex) => (modelMap, rowIndex == repeatedData.size - 1) }

    lazy val writeIndex = dataForRow.foldLeft(writeRowIndex) { case (currWriteIndex, (modelMap, isLastDataMap)) =>
      repeatedRowsWithPartialWrittenData.add(templateIndex)
      createRow(templateRow, modelMap, currWriteIndex)
      lazy val shouldSkipNextRow = templateSheet.rowOpt(templateRow.getRowNum + 1)
        .forall(row => shouldSkipRow(row, staticData))
      if (isLastDataMap && shouldSkipNextRow) {
        return RowWriteResult(currWriteIndex + 1, shouldStopCurrentWrite = true)
      }
      else {
        currWriteIndex + 1
      }
    }

    if (dataForRow.nonEmpty) {
      RowWriteResult(writeIndex, shouldStopCurrentWrite = dataForRow.isEmpty)
    }
    else if (repeatedRowsWithPartialWrittenData.contains(templateRow.getRowNum)) {
      templateIndex += 1
      RowWriteResult(writeRowIndex, shouldStopCurrentWrite = false)
    }
    else {
      RowWriteResult(writeRowIndex, shouldStopCurrentWrite = true)
    }
  }

  private def initialWrite(): Unit = {
    copyPictures()
    templateSheet.rowIterator().asScala
      .toList
      .headOption
      .filter(row => !row.isRepeatedRow && !shouldSkipRow(row, Map.empty))
      .foreach { templateRow =>
        createRow(templateRow, Map.empty, outputRowIndex)
        outputRowIndex += 1
        templateIndex += 1
      }
  }

  private def isSubstituteCell(cell: Cell): Boolean = {
    cell.getCellType == CellType.STRING &&
      (cell.getStringCellValue.startsWith(ExportModelUtils.REPEATED_FIELD_KEY) ||
        cell.getStringCellValue.startsWith(ExportModelUtils.SUBSTITUTION_KEY))
  }

  private def createRow(templateRow: Row, modelMap: ModelMap, writeIndex: Int): Unit = {
    val outputRow = outputSheet.createRow(writeIndex)
    populateRowFromTemplate(templateRow, outputRow, templateSheet, outputSheet, modelMap)
    numRowsCreated += 1
  }

  private def populateRowFromTemplate(templateRow: Row,
                                      outputRow: Row,
                                      templateSheet: Sheet,
                                      outputSheet: Sheet,
                                      substitutionMap: ModelMap
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

  private def substituteAndCopyCell(templateCell: Cell, outputCell: Cell, substitutionMap: ModelMap): Unit = {
    templateCell.getCellType match {
      case CellType.NUMERIC =>
        outputCell.setCellValue(templateCell.getNumericCellValue)
      case CellType.BOOLEAN =>
        outputCell.setCellValue(templateCell.getBooleanCellValue)
      case CellType.STRING if isSubstituteCell(templateCell) =>
        substituteString(templateCell, outputCell, substitutionMap)
      case CellType.STRING =>
        outputCell.setCellValue(templateCell.getStringCellValue)
      case CellType.FORMULA =>
        val cellRefShiftFactor = numRowsCreated - templateCell.getRowIndex
        val outputFormula = cellFormulaParser.shiftRowNums(templateCell.getCellFormula, cellRefShiftFactor)
        outputCell.setCellFormula(outputFormula)
      case _ =>
      //logger.trace("CellType is one of {CellType.ERROR, CellType.BLANK, CellType.FORMULA, CellType.NONE} which we do not want to copy from")
    }
  }

  private def substituteString(templateCell: Cell, outputCell: Cell, substitutionMap: ModelMap): Unit = {
    val templateValue = templateCell.getStringCellValue
    substitutionMap.getOrElse(templateValue, CellBlank()) match {
      case CellDouble(double: Double) => outputCell.setCellValue(double)
      case CellDate(date: Date) => outputCell.setCellValue(date)
      case CellString(string: String) => outputCell.setCellValue(string)
      case CellBlank() => outputCell.setBlank()
      case _ => throw new NoSuchElementException(s"Unsupported cell type for key: $templateValue")
    }
  }

  private def shouldSkipRow(templateRow: Row, modelMaps: ModelMap*): Boolean = {
    val substituteCells = templateRow.cellIterator().asScala
      .filter(isSubstituteCell)
      .toList

    if (substituteCells.isEmpty) {
      false
    }
    else {
      !substituteCells.map(_.getStringCellValue)
        .exists(cellValue => modelMaps.exists(modelMap => modelMap.contains(cellValue)))
    }
  }
}

object AppendOnlySheetWriter {
  def apply(templateSheet: XSSFSheet, outputWorkbook: SXSSFWorkbook, schema: Map[String, YamlEntry], cellStyleCache: mutable.Map[CellStyle, Int]): AppendOnlySheetWriter = {
    val sheetWriter = new AppendOnlySheetWriter(templateSheet, outputWorkbook, schema, cellStyleCache)
    sheetWriter.initialWrite()
    sheetWriter
  }
}