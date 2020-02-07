package com.carta.excel

import java.io.{Closeable, InputStream, OutputStream}
import java.util.Date

import com.carta.excel.ExportModelUtils.ModelMap
import org.apache.poi.ss.usermodel.{CellStyle, CellType}
import org.apache.poi.xssf.streaming.{SXSSFCell, SXSSFRow, SXSSFSheet, SXSSFWorkbook}
import org.apache.poi.xssf.usermodel.{XSSFCell, XSSFRow, XSSFSheet, XSSFWorkbook}

import scala.collection.{immutable, mutable}
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
  private val excelTemplateWorkbooks: Map[String, XSSFWorkbook] = templateStreamMap
    .mapValues(streamResource => streamResource
      .acquireAndGet(stream => new XSSFWorkbook(stream)))
  private val outputExcelWorkbook: SXSSFWorkbook = new SXSSFWorkbook(windowSize)

  // Cache of cell styles to avoid duplicating copied cell styles across various templates in the output workbook.
  private val cellStyleCache: mutable.HashMap[CellStyle, Int] = new mutable.HashMap[CellStyle, Int]

  // copyAndSubstitute copies the template workbook given by templateName to the output workbook in a
  // buffered streamed manner, using substitutionMap to make any key value substitutions that are encountered
  // during copying.
  //
  // Returns the current index within the template to start adding rows as expected by the template or
  // None if this is the end of this template and no rows are expected to be inserted.
  def copyAndSubstitute(templateName: String, substitutionMap: ExportModelUtils.ModelMap = Map.empty): immutable.Seq[(String, Option[Int])] = {
    // We expect every template workbook to only have 1 sheet - this allows us to be more
    // flexible with what templates we want to load into memory via the XSSF api.
    val templateWorkbook: XSSFWorkbook = excelTemplateWorkbooks(templateName)
    val templateSheets: immutable.Seq[XSSFSheet] = (0 until templateWorkbook.getNumberOfSheets())
        .map(templateWorkbook.getSheetAt(_))

    templateSheets.map { templateSheet =>
      val sheetName = templateSheet.getSheetName()
      // Create sheet in workbook to write to
      val outputSheet: SXSSFSheet = outputExcelWorkbook.createSheet(sheetName)

      copyAndSubstitute(templateSheet, outputSheet, substitutionMap, sheetName)
    }
  }

  private def copyAndSubstitute(templateSheet: XSSFSheet, outputSheet: SXSSFSheet, substitutionMap: ModelMap, sheetName: String): (String, Option[Int]) = {
    // Get each template row and its corresponding row index and use the given substitutionMap to copy them
    // over to the output workbook. If we find a placeholder repeated row in the template, then we return that
    // row's index to be returned to the caller who will use it to start inserting repeated rows.
    // We need to explicitly break and return at this point in order to ensure that no further rows are written
    // and the stream's pointer is not advanced.
    var x: (String, Option[Int]) = (sheetName, None)
    for (rowIndex <- 0 to templateSheet.getLastRowNum()) {
      Option(templateSheet.getRow(rowIndex)) match {
        // A row may be empty in the template
        case Some(templateRow) =>
          if (!isRepeatedRow(templateRow)) {
            val outputRow = outputSheet.createRow(rowIndex)
            populateRowFromTemplate(templateRow, outputRow, templateSheet, outputSheet, substitutionMap)
          } else {
            // Do not copy and substitute anything from the template lower than a repeated row
            x = (sheetName, Some(rowIndex))
          }
        case None =>
      }
    }

    x
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
  def insertRows(templateName: String, templateRowIndex: Int, tabName: String, outputStartIndex: Int, rows: immutable.Seq[ExportModelUtils.ModelMap]): Try[Int] = {
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
  def write(os: OutputStream): Try[Unit] = {
    Try(outputExcelWorkbook.write(os))
  }

  // close releases and closes all underlying resources of this ExcelWorkbook, namely the input
  // template workbooks and the output write workbook.
  def close(): Unit = {
    excelTemplateWorkbooks.foreach {
      case (_, wb) => wb.close()
    }

    // dispose clears temporary files that back the workbook on disk
    outputExcelWorkbook.dispose()
    outputExcelWorkbook.close()
  }

  private def isRepeatedRow(row: XSSFRow): Boolean = {
    row.getLastCellNum() > 0 &&
      Option(row.getCell(0)) != None &&
      row.getCell(0).getCellType() == CellType.STRING &&
      row.getCell(0).getStringCellValue().startsWith(ExportModelUtils.REPEATED_FIELD_KEY)
  }

  private def populateRowFromTemplate(templateRow: XSSFRow,
                                       outputRow: SXSSFRow,
                                       templateSheet: XSSFSheet,
                                       outputSheet: SXSSFSheet,
                                       substitutionMap: ExportModelUtils.ModelMap
                                     ): Unit = {
    copyRowHeight(templateRow, outputRow)

    for (colIndex <- 0 until templateRow.getLastCellNum()) {
      Option(templateRow.getCell(colIndex)) match {
        case Some(templateCell: XSSFCell) =>
          val outputCell = outputRow.createCell(colIndex)

          copyColumnWidth(templateSheet, outputSheet, colIndex)
          applyCellStyleFromTemplate(templateCell, outputCell)
          substituteAndCopyCell(templateCell, outputCell, substitutionMap)
        case None => Unit
      }
    }
  }

  private def substituteAndCopyCell(templateCell: XSSFCell, outputCell: SXSSFCell, substitutionMap: ExportModelUtils.ModelMap): Unit = {
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

  private def substituteString(stringValue: String, outputCell: SXSSFCell, substitutionMap: ExportModelUtils.ModelMap): Unit = {
    if (stringValue.startsWith(ExportModelUtils.SUBSTITUTION_KEY) || stringValue.startsWith(ExportModelUtils.REPEATED_FIELD_KEY)){
      if (substitutionMap.contains(stringValue)) {
        substitutionMap(stringValue) match {
          case CellDouble(double: Double) => outputCell.setCellValue(double)
          case CellDate(date: Date) => outputCell.setCellValue(date)
          case CellString(string: String) => outputCell.setCellValue(string)
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

  private def applyCellStyleFromTemplate(fromCell: XSSFCell, toCell: SXSSFCell): Unit = {
    val templateStyle: CellStyle = fromCell.getCellStyle()

    if (!cellStyleCache.contains(templateStyle)) {
      val outputStyle = outputExcelWorkbook.createCellStyle()

      outputStyle.cloneStyleFrom(templateStyle)
      toCell.setCellStyle(outputStyle)

      cellStyleCache.put(templateStyle, outputStyle.getIndex())
    }
    else {
      toCell.setCellStyle(outputExcelWorkbook.getCellStyleAt(cellStyleCache(templateStyle)))
    }
  }

  private def copyColumnWidth(from: XSSFSheet, to: SXSSFSheet, col: Int): Unit = {
    to.setColumnWidth(col, from.getColumnWidth(col))
  }

  private def copyRowHeight(from: XSSFRow, to: SXSSFRow): Unit = {
    to.setHeight(from.getHeight)
  }
}
