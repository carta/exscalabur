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

import com.fasterxml.jackson.core.`type`.TypeReference
import com.fasterxml.jackson.databind.ObjectMapper
import com.fasterxml.jackson.dataformat.yaml.YAMLFactory
import com.fasterxml.jackson.module.scala.DefaultScalaModule

import scala.io.Source

class YamlReader(private val mapper: ObjectMapper) {
  private val parseTypeRef = new TypeReference[Map[String, EntryBuilder]]() {}

  def parse(file: File): Map[String, YamlEntry] = {
    val dataSource = Source.fromFile(file)
    val yamlData = dataSource.mkString
    dataSource.close()
    parseFromContent(yamlData)
  }

  def parse(path: String): Map[String, YamlEntry] = {
    this.parse(new File(path))
  }

  private def parseFromContent(content: String): Map[String, YamlEntry] = {
    if (content.isEmpty) {
      Map.empty
    }
    else {
      val yamlMap: Map[String, EntryBuilder] = mapper.readValue(content, parseTypeRef)
      yamlMap.mapValues(entryBuilder => entryBuilder.build())
    }
  }
}

object YamlReader {
  def apply(): YamlReader = {
    val mapper = new ObjectMapper(new YAMLFactory()).registerModule(DefaultScalaModule)
    mapper.registerModule(DefaultScalaModule)
    new YamlReader(mapper)
  }
}
