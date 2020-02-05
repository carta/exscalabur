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
package com.carta.excel

import java.util.Date

sealed trait CellValue

final case class CellString(s: String) extends CellValue

final case class CellDouble(d: Double) extends CellValue

final case class CellDate(d: Date) extends CellValue

final case class CellBoolean(b: Boolean) extends CellValue

final case class CellBlank() extends CellValue