package com.carta.exscalabur

import java.io.{FileOutputStream, OutputStream}

import com.carta.excel.AppendOnlySheetWriter
import com.carta.yaml.YamlEntry
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.{XSSFSheet, XSSFWorkbook}

import scala.collection.mutable

//TODO verify no keys in yaml match
class Exscalabur(sheetsByName: Map[String, XSSFSheet],
                 outputStream: OutputStream,
                 outputWorkbook: SXSSFWorkbook,
                 cellStyleCache: mutable.Map[CellStyle, Int],
                 schema: Map[String, YamlEntry]) {

  def getAppendOnlySheetWriter(sheetName: String): AppendOnlySheetWriter = {
    val sheet = sheetsByName(sheetName)
    AppendOnlySheetWriter(sheet, outputWorkbook, schema, cellStyleCache)
  }

  def writeToDisk(): Unit = {
    outputWorkbook.write(outputStream)
    outputWorkbook.dispose()
    outputWorkbook.close()
    outputStream.close()
  }
}

object Exscalabur {
  def apply(templatePaths: Iterable[String],
            outputPath: String,
            schema: Map[String, YamlEntry],
            windowSize: Int): Exscalabur = {
    val sheetsByName = templatePaths.map { str =>
      val sheet = new XSSFWorkbook(str).getSheetAt(0)
      sheet.getSheetName -> sheet
    }.toMap
    val outputStream = new FileOutputStream(outputPath)

    val outputWorkbook = new SXSSFWorkbook(windowSize)
    val cellStyleCache = mutable.Map.empty[CellStyle, Int]

    new Exscalabur(sheetsByName, outputStream, outputWorkbook, cellStyleCache, schema)
  }
}