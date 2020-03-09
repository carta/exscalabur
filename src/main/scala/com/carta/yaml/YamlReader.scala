package com.carta.yaml

import java.io.File

import com.fasterxml.jackson.core.`type`.TypeReference
import com.fasterxml.jackson.databind.ObjectMapper
import com.fasterxml.jackson.dataformat.yaml.YAMLFactory
import com.fasterxml.jackson.module.scala.DefaultScalaModule

import scala.io.Source

class YamlReader(private val mapper: ObjectMapper) {
  private val parseTypeRef = new TypeReference[Map[String, KeyObjectBuilder]]() {}

  def parse(file: File): Map[String, YamlEntry] = {
    val dataSource = Source.fromFile(file)
    val yamlData = dataSource.mkString
    dataSource.close()
    parseFromContent(yamlData)
  }

  def parse(resourcePath: String): Map[String, YamlEntry] = {
    val yamlData = Source.fromResource(resourcePath).mkString
    parseFromContent(yamlData)
  }

  private def parseFromContent(content: String): Map[String, YamlEntry] = {
    if (content.isEmpty) {
      Map.empty
    }
    else {
      val yamlMap: Map[String, KeyObjectBuilder] = mapper.readValue(content, parseTypeRef)
      yamlMap.map { case (yamlKey, keyObject) =>
        yamlKey -> keyObject.build()
      }
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
