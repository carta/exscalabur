package com.carta.compat.java.excel


import com.carta.compat.java.exscalabur.{DataCell, DataRow}

import scala.collection.JavaConverters._

class AppendOnlySheetWriter(scalaSheetWriter: com.carta.excel.AppendOnlySheetWriter) {
  def writeStaticData(data: java.util.List[DataCell]): Unit = {
    scalaSheetWriter.writeStaticData(data.asScala.map(_.asScala))
  }

  def writeRepeatedData(dataRows: java.util.List[DataRow]): Unit = {
    scalaSheetWriter.writeRepeatedData(dataRows.asScala.map(_.asScala))
  }

  def copyPictures(): Unit = {
    scalaSheetWriter.copyPictures()
  }
}
