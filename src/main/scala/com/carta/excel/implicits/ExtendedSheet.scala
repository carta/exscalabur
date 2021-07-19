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
      .toSeq
  }

}

