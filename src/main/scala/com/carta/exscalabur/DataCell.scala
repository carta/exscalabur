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

case class DataCell(key: String, value: CellType) {
  def asTuple: (String, CellType) = key -> value
}

object DataCell {
  def apply(key: String, value: String): DataCell = apply(key, StringCellType(value))
  def apply(key: String, value: Long): DataCell = apply(key, LongCellType(value))
  def apply(key: String, value: Double): DataCell = apply(key, DoubleCellType(value))
}
