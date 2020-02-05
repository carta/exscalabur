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

import com.carta.excel.ExportModelUtils
import org.apache.poi.ss.usermodel.{Cell, CellType, Row}

import scala.collection.JavaConverters._

object ExtendedRow {

  implicit class ExtendedRow(row: Row) {
    def isRepeatedRow: Boolean = rowHasCellStartingWithKey(ExportModelUtils.REPEATED_FIELD_KEY)

    def isStaticRow: Boolean = rowHasCellStartingWithKey(ExportModelUtils.SUBSTITUTION_KEY)

    private def rowHasCellStartingWithKey(key: String): Boolean = {
      row.cellIterator().asScala
        .filter(cell => cell.getCellType == CellType.STRING)
        .map(cell => cell.getStringCellValue)
        .exists(cellValue => cellValue.startsWith(key))
    }
  }

}
