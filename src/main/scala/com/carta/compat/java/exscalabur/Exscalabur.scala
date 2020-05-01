package com.carta.compat.java.exscalabur

import java.io.OutputStream

import com.carta.compat.java.excel.AppendOnlySheetWriter
import com.carta.yaml.YamlEntry

import scala.collection.JavaConverters._

class Exscalabur(scalaExscalabur: com.carta.exscalabur.Exscalabur) {
  def this(templatePaths: java.util.List[String],
           schema: java.util.Map[String, YamlEntry],
           windowSize: Int) = {
    this(
      scalaExscalabur = com.carta.exscalabur.Exscalabur(templatePaths.asScala, schema.asScala.toMap, windowSize)
    )
  }

  def getAppendOnlySheetWriter(sheetName: String): AppendOnlySheetWriter = {
    val scalaSheetWriter = scalaExscalabur.getAppendOnlySheetWriter(sheetName)
    new com.carta.compat.java.excel.AppendOnlySheetWriter(scalaSheetWriter)
  }

  def exportToFile(path: String): Unit = {
    scalaExscalabur.exportToFile(path)
  }

  def exportToStream(outputStream: OutputStream): Unit = {
    scalaExscalabur.exportToStream(outputStream)
  }
}
