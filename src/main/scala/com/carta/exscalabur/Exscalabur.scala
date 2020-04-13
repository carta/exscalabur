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