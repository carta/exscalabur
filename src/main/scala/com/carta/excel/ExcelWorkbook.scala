package com.carta.excel

import java.io.{Closeable, InputStream, OutputStream}
import java.util.Date

import com.carta.excel.ExportModelUtils.ModelMap
import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.streaming.{SXSSFSheet, SXSSFWorkbook}
import org.apache.poi.xssf.usermodel.{XSSFPicture, XSSFSheet, XSSFWorkbook}

import scala.collection.JavaConverters._
import scala.collection.mutable
import scala.util.{Failure, Success, Try}

/**
 * ExcelWorkbook is an excel generation class that wraps the Java Apache POI library. Its inputs are a set of template workbooks to read from
 * and the buffer window size which is used to stream writes to the output workbook.
 *
 * For information on modifying formatting, styling, or contents of the final output workbook, see the wiki here:
 * https://github.com/carta/ds-reporting-service/wiki/Exports
 *
 * @param windowSize the number of rows that are kept in memory until flushed out, -1 means all records are available for random access .
 */
class ExcelWorkbook(templateStreamMap: Map[String, resource.ManagedResource[InputStream]], windowSize: Int) extends Closeable {
  // We use an InputStream instead of a File when creating the XSSFWorkbook to ensure that the read file is not
  // written to at all. When using a File, XSSFWorkbook#close() will actually modify the last access timestamp of
  // the file unnecessarily and become annoying for git diffs. Using InputStream here ensures that this does
  // not happen (setting the file as readonly is not enough, XSSFWorkbook#close() will still try to write and actually
  // raise an exception.
  // Apache POI documentation warns against using the InputStream because it actually has a bigger memory footprint
  // than using an XSSFWorkbook with a File, but for our purposes which are just to open and copy small templates, this
  // shouldn't be a problem.

  private val outputExcelWorkbook: SXSSFWorkbook = new SXSSFWorkbook(windowSize)

  private val excelTemplateWorkbooks: Map[String, XSSFWorkbook] = {
    templateStreamMap.mapValues { streamResource =>
      streamResource.acquireAndGet(stream => new XSSFWorkbook(stream))
    }
  }

  def sheets(templateName: String): Seq[Sheet] = {
    excelTemplateWorkbooks(templateName).sheetIterator()
      .asScala
      .toVector
  }

  // Cache of cell styles to avoid duplicating copied cell styles across various templates in the output workbook.
  private val cellStyleCache: mutable.HashMap[CellStyle, Int] = new mutable.HashMap[CellStyle, Int]

  // insertRows inserts the given rows into the given index position in the template
  // indicated by the given templateName.
  //
  // Returns the index position to insert the next Seq of rows to the template or None if there are
  // no more rows expected to be inserted.
  /**
   * @param templateName
   * @param templateRowIndex index at which to get the style from the template sheet
   * @param outputStartIndex index at where to start writing outputs
   * @param rows
   * @return
   */
  def insertRows(templateName: String, templateRowIndex: Int, sheetName: String, outputStartIndex: Int, staticData: ModelMap, rows: Seq[ModelMap]): Try[Int] = {
    if (templateRowIndex < 0 || outputStartIndex < 0) {
      Failure(new IllegalArgumentException(s"Invalid indices - templateRowIndex=$templateRowIndex, outputStartIndex=$outputStartIndex"))
    }
    else if (!excelTemplateWorkbooks.contains(templateName)) {
      Failure(new IllegalArgumentException(s"No valid template workbook found for template name: ${templateName}"))
    }
    else {
      val templateSheet = excelTemplateWorkbooks(templateName).getSheet(sheetName)
      if (Option(outputExcelWorkbook.getSheet(sheetName)).isEmpty) {
        outputExcelWorkbook.createSheet(sheetName)
      }
      val outputSheet = outputExcelWorkbook.getSheet(sheetName)
      var nextIndex = outputStartIndex
      Option(templateSheet.getRow(templateRowIndex)).foreach { templateRow =>
        if (isRepeatedRow(templateRow)) {
          rows.foreach{ modelMap =>
            val outputRow = outputSheet.createRow(nextIndex)
            populateRowFromTemplate(templateRow, outputRow, templateSheet, outputSheet, modelMap)
            nextIndex += 1
          }
        }
        else {
          val outputRow = outputSheet.createRow(nextIndex)
          populateRowFromTemplate(templateRow, outputRow, templateSheet, outputSheet, staticData)
          nextIndex += 1
        }
      }
      copyPicturesToSheet(templateSheet, outputSheet)
      Success(nextIndex)
    }
  }


  // write writes the output workbook to the given OutputStream
  def write(os: OutputStream): Try[Unit] = Try {
    outputExcelWorkbook.write(os)
  }

  // close releases and closes all underlying resources of this ExcelWorkbook, namely the input
  // template workbooks and the output write workbook.
  def close(): Unit = {
    excelTemplateWorkbooks.values.foreach(_.close())

    // dispose clears temporary files that back the workbook on disk
    outputExcelWorkbook.dispose()
    outputExcelWorkbook.close()
  }

  //TODO if any cell on row is repeated cell, return true
  private def isRepeatedRow(row: Row): Boolean = {
    row.getLastCellNum > 0 && isRepeatedCell(Option(row.getCell(0)))
  }

  private def isRepeatedCell(cellOpt: Option[Cell]): Boolean = cellOpt match {
    case None => false
    case Some(cell) =>
      CellType.STRING == cell.getCellType &&
        cell.getStringCellValue.startsWith(ExportModelUtils.REPEATED_FIELD_KEY)
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

  private def substituteAndCopyCell(templateCell: Cell, outputCell: Cell, substitutionMap: ExportModelUtils.ModelMap): Unit = {
    templateCell.getCellType match {
      case CellType.NUMERIC =>
        outputCell.setCellValue(templateCell.getNumericCellValue)
      case CellType.BOOLEAN =>
        outputCell.setCellValue(templateCell.getBooleanCellValue)
      case CellType.STRING =>
        val stringValue = templateCell.getStringCellValue
        substituteString(stringValue, outputCell, substitutionMap)
      case _ =>
      //logger.trace("CellType is one of {CellType.ERROR, CellType.BLANK, CellType.FORMULA, CellType.NONE} which we do not want to copy from")
    }
  }

  private def substituteString(stringValue: String, outputCell: Cell, substitutionMap: ExportModelUtils.ModelMap): Unit = {
    if (stringValue.startsWith(ExportModelUtils.SUBSTITUTION_KEY) || stringValue.startsWith(ExportModelUtils.REPEATED_FIELD_KEY)) {
      if (substitutionMap.contains(stringValue)) {
        substitutionMap(stringValue) match {
          case CellDouble(double: Double) => outputCell.setCellValue(double)
          case CellDate(date: Date) => outputCell.setCellValue(date)
          case CellString(string: String) => outputCell.setCellValue(string)
          case CellBoolean(bool: Boolean) => outputCell.setCellValue(bool)
          case _ => throw new NoSuchElementException(s"Unsupported cell type for key: $stringValue")
        }
      }
      else {
        //logger.info(s"Key from template is not defined in substitution Map: $stringValue")
      }
    }
    else {
      outputCell.setCellValue(stringValue)
    }
  }

  private def copyPicturesToSheet(templateSheet: XSSFSheet, outputSheet: Sheet): Unit = {
    val outputWorkbook = outputSheet.getWorkbook
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

  private def applyCellStyleFromTemplate(from: Cell, to: Cell): Unit = {
    val templateStyle = from.getCellStyle

    if (!cellStyleCache.contains(templateStyle)) {
      val outputStyle = outputExcelWorkbook.createCellStyle()
      outputStyle.cloneStyleFrom(templateStyle)
      cellStyleCache.put(templateStyle, outputStyle.getIndex)
    }

    val outputStyleIndex = cellStyleCache(templateStyle)
    to.setCellStyle(outputExcelWorkbook.getCellStyleAt(outputStyleIndex))
  }

  private def copyColumnWidth(from: Sheet, to: Sheet, col: Int): Unit = {
    to.setColumnWidth(col, from.getColumnWidth(col))
  }

  private def copyRowHeight(from: Row, to: Row): Unit = {
    to.setHeight(from.getHeight)
  }
}
