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
package com.carta.yaml

import java.io.File

import com.carta.UnitSpec

class YamlReaderSpec extends UnitSpec {
  val testDataPath: String = getClass.getResource("/test.yaml").getFile
  val testDataFile: File = new File(testDataPath)

  "YAML Reader" should "Produce YamlEntries given yaml file" in {
    val expectedData = Vector(
      YamlEntry(DataType.string, ExcelType.string),
      YamlEntry(DataType.double, ExcelType.number),
      YamlEntry(DataType.long, ExcelType.date)
    )

    val yamlReader = YamlReader()
    val actualData = yamlReader.parse(testDataPath).values

    actualData should contain theSameElementsAs expectedData
  }

  it should "Produce KeyObject from yaml file as File" in {
    val keyObjects = YamlReader().parse(testDataFile)

    val expectedData = Vector(
      YamlEntry(DataType.string, ExcelType.string),
      YamlEntry(DataType.double, ExcelType.number),
      YamlEntry(DataType.long, ExcelType.date)
    )

    keyObjects.values should contain theSameElementsAs expectedData
  }
}