package com.carta.excel.implicits

import org.apache.poi.ss.usermodel.{ClientAnchor, Row, Sheet}
import org.apache.poi.xssf.usermodel.{XSSFPicture, XSSFSheet}

import scala.collection.JavaConverters._

object ExtendedSheet {

  implicit class ExtendedSheet(sheet: Sheet) {
    def getRowIndices: Iterable[Int] = (0 to sheet.getLastRowNum).toVector

    def rowOpt(rowIndex: Int): Option[Row] = Option(sheet.getRow(rowIndex))

    def copyPicture(img: XSSFPicture): Unit = {
      val drawing = sheet.createDrawingPatriarch()
      val anchor = sheet.getWorkbook.getCreationHelper.createClientAnchor()
      anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE)
      val imgData = img.getPictureData
      val templateAnchor = img.getClientAnchor
      val picIndex = sheet.getWorkbook.addPicture(imgData.getData, imgData.getPictureType)

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

  implicit class ExtendedXSSFSheet(sheet: XSSFSheet) extends ExtendedSheet(sheet) {
    def getPictures: Seq[XSSFPicture] = sheet.createDrawingPatriarch()
      .getShapes.asScala
      .filter(_.isInstanceOf[XSSFPicture])
      .map(_.asInstanceOf[XSSFPicture])
  }

}

