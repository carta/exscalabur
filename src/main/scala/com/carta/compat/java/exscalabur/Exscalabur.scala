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
