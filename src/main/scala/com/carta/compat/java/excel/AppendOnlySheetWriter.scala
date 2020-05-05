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
