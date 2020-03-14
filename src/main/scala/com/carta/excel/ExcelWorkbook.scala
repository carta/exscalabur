package com.carta.excel

import java.io.{Closeable, InputStream, OutputStream}
import java.util.Date

import com.carta.excel.ExportModelUtils.ModelMap
import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.streaming.SXSSFWorkbook
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

  private val cellStyleCache = mutable.Map.empty[CellStyle, Int]


  /** Copies a row from the template worksheet under templateName to the output sheet (creating one if none exist).
   *
   * If the to-be-copied row is a "static" row, data from staticData will be substituted into cells with matching keys,
   * otherwise, will be copied as is. In this case, the returned value will be one more than the outputStartIndex
   *
   * If the to-be-copied row is a "repeated" row, data from repeatedRows will be substituted into cells with matching keys.
   * A new row in the output sheet will be created for each data element found in repeatedData for that key. In which case,
   * the returned value will be outputStartIndex + repeatedData.length
   *
   * @param templateName     sheet from template workbook to copy data from, and write data to in output workbook
   * @param templateRowIndex current index of template to write to
   * @param outputStartIndex current index of output sheet to write to
   * @param staticData       Map of static data keys to their corresponding values
   * @param repeatedData     List of Maps from repeated data keys to their corresponding values
   * @return the next row in the output sheet to write to
   */
  def insertRows(templateName: String,
                 templateRowIndex: Int,
                 sheetName: String,
                 outputStartIndex: Int,
                 staticData: ModelMap,
                 repeatedData: Seq[ModelMap]): Try[Int] = {
    if (templateRowIndex < 0 || outputStartIndex < 0) {
      Failure(new IllegalArgumentException(s"Invalid indices - templateRowIndex=$templateRowIndex, outputStartIndex=$outputStartIndex"))
    }
    else if (!excelTemplateWorkbooks.contains(templateName)) {
      Failure(new IllegalArgumentException(s"No valid template workbook found for template name: $templateName"))
    }
    else {
      val templateSheet = excelTemplateWorkbooks(templateName).getSheet(sheetName)
      val outputSheet = Option(outputExcelWorkbook.getSheet(sheetName))
        .getOrElse(outputExcelWorkbook.createSheet(sheetName))

      val nextIndex = Option(templateSheet.getRow(templateRowIndex)) match {
        case Some(templateRow) =>
          if (isRepeatedRow(templateRow)) {
            repeatedData.foldLeft(outputStartIndex) { (currWriteIndex, modelMap) =>
              if (shouldSkipRow(templateRow, modelMap)) {
                currWriteIndex
              }
              else {
                createRow(outputSheet, templateSheet, templateRow, modelMap, currWriteIndex)
                currWriteIndex + 1
              }
            }
          }
          else {
            createRow(outputSheet, templateSheet, templateRow, staticData, outputStartIndex)
            outputStartIndex + 1
          }
        case None =>
          createBlankRow(outputSheet, outputStartIndex)
          outputStartIndex + 1
      }

      copyPicturesToSheet(templateSheet, outputSheet)
      Success(nextIndex)
    }
  }

  private def shouldSkipRow(templateRow: Row, modelMap: ModelMap): Boolean = {
    println(templateRow.getLastCellNum)
    templateRow.cellIterator().asScala.exists { cell =>
      if (cell.getCellType != CellType.STRING) {
        false
      }
      else if (!cell.getStringCellValue.startsWith(ExportModelUtils.REPEATED_FIELD_KEY)) {
        false
      }
      else {
        !modelMap.contains(cell.getStringCellValue)
      }
    }
  }

  private def createBlankRow(outputSheet: Sheet, index: Int): Unit = {
    outputSheet.createRow(index)
  }

  private def createRow(outputSheet: Sheet, templateSheet: Sheet, templateRow: Row, modelMap: ModelMap, writeIndex: Int): Unit = {
    val outputRow = outputSheet.createRow(writeIndex)
    populateRowFromTemplate(templateRow, outputRow, templateSheet, outputSheet, modelMap)
  }

  def sheets(templateName: String): Seq[Sheet] = {
    excelTemplateWorkbooks(templateName).sheetIterator()
      .asScala
      .toVector
  }


  // write writes the output workbook to the given OutputStream
  def write(os: OutputStream): Try[Unit] = Try {
    outputExcelWorkbook.write(os)
  }

  // close releases and closes all underlying resources of this ExcelWorkbook, namely the input
  // template workbooks and the output write workbook.
  override def close(): Unit = {
    excelTemplateWorkbooks.values.foreach(_.close())

    // dispose clears temporary files that back the workbook on disk
    outputExcelWorkbook.dispose()
    outputExcelWorkbook.close()
  }

  private def isRepeatedRow(row: Row): Boolean = {
    row.cellIterator().asScala.exists(isRepeatedCell)
  }

  private def isRepeatedCell(cell: Cell): Boolean = {
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
