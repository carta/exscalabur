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


  private val excelTemplateWorkbooks: Map[String, XSSFWorkbook] = {
    templateStreamMap.mapValues { streamResource =>
      streamResource.acquireAndGet(stream => new XSSFWorkbook(stream))
    }
  }
  private val outputExcelWorkbook: SXSSFWorkbook = new SXSSFWorkbook(windowSize)

  // Cache of cell styles to avoid duplicating copied cell styles across various templates in the output workbook.
  private val cellStyleCache: mutable.HashMap[CellStyle, Int] = new mutable.HashMap[CellStyle, Int]

  // copyAndSubstitute copies the template workbook given by templateName to the output workbook in a
  // buffered streamed manner, using substitutionMap to make any key value substitutions that are encountered
  // during copying.
  //
  // Returns the current index within the template to start adding rows as expected by the template or
  // None if this is the end of this template and no rows are expected to be inserted.
  def copyAndSubstitute(templateName: String, substitutionMap: ExportModelUtils.ModelMap = Map.empty): Iterable[(String, Option[Int])] = {
    // We expect every template workbook to only have 1 sheet - this allows us to be more
    // flexible with what templates we want to load into memory via the XSSF api.
    excelTemplateWorkbooks(templateName).sheetIterator()
      .asScala
      .toVector
      .map(_.asInstanceOf[XSSFSheet])
      .map { templateSheet =>
        val sheetName = templateSheet.getSheetName
        val outputSheet = outputExcelWorkbook.createSheet(sheetName)

        val index = substituteStaticRows(templateSheet, outputSheet, substitutionMap)

        val numColumns = getNumColumns(templateSheet)
        (0 until numColumns).foreach { i =>
          outputSheet.setColumnWidth(i, templateSheet.getColumnWidth(i))
        }
        copyPicturesToSheet(templateSheet, outputSheet)
        (sheetName, index)
      }
  }

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
  def insertRows(templateName: String, templateRowIndex: Int, tabName: String, outputStartIndex: Int, rows: Seq[ExportModelUtils.ModelMap]): Try[Int] = {
    // Indices are invalid
    if (templateRowIndex < 0 || outputStartIndex < 0) {
      Failure(new IllegalArgumentException(s"Invalid Indices templateRowIndex=$templateRowIndex outputStartIndex=$outputStartIndex"))
    }
    // Template name does not correspond to a valid template workbook
    else if (!excelTemplateWorkbooks.contains(templateName)) {
      //logger.warn(s"No valid template workbook found for template name: ${templateName}")
      Failure(new IllegalArgumentException(s"No valid template workbook found for template name: ${templateName}"))
    }
    else {
      val templateSheet: XSSFSheet = excelTemplateWorkbooks(templateName).getSheet(tabName)
      val outputSheet: SXSSFSheet = outputExcelWorkbook.getSheet(tabName)

      Option(templateSheet.getRow(templateRowIndex)) match {
        case Some(templateRepeatedRow) =>
          for (rowIndex <- 0 until rows.length) {
            try {
              val outputRow = outputSheet.createRow(outputStartIndex + rowIndex)
              populateRowFromTemplate(templateRepeatedRow, outputRow, templateSheet, outputSheet, rows(rowIndex))
            } catch {
              case e: IllegalArgumentException =>
              //logger.warn(s"Failed to create row at index ${outputStartIndex + rowIndex} in output workbook because row is already created")
            }
          }
          Success(outputStartIndex + rows.length)

        case None =>
          //logger.warn(s"No valid template row to base repeated rows on found for template ${templateName} at index ${templateRowIndex}")
          Failure(new NoSuchElementException(s"No valid template row to base repeated rows on found for template ${templateName} at index ${templateRowIndex}"))
      }
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

  private def substituteStaticRows(templateSheet: Sheet, outputSheet: SXSSFSheet, substitutionMap: ModelMap): Option[Int] = {
    // Get each template row and its corresponding row index and use the given substitutionMap to copy them
    // over to the output workbook. If we find a placeholder repeated row in the template, then we return that
    // row's index to be returned to the caller who will use it to start inserting repeated rows.
    // We need to explicitly break and return at this point in order to ensure that no further rows are written
    // and the stream's pointer is not advanced.
    templateSheet.iterator().asScala
      .foldLeft(Option.empty[Int]) { (acc: Option[Int], templateRow: Row) =>
        val rowIndex = templateRow.getRowNum
        if (!isRepeatedRow(templateRow)) {
          val outputRow = outputSheet.createRow(rowIndex)
          populateRowFromTemplate(templateRow, outputRow, templateSheet, outputSheet, substitutionMap)
          acc
        } else {
          // only copies one repeated section at a time
          Some(rowIndex)
        }
      }
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


  private def getNumColumns(sheet: Sheet) = {
    sheet.rowIterator().asScala
      .map(row => row.getLastCellNum.toInt)
      .foldRight(0) {
        Math.max
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
