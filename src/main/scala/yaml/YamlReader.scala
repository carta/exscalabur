package yaml

import java.io.File

import com.fasterxml.jackson.core.`type`.TypeReference
import com.fasterxml.jackson.databind.ObjectMapper
import com.fasterxml.jackson.dataformat.yaml.YAMLFactory
import com.fasterxml.jackson.module.scala.DefaultScalaModule

import scala.io.Source

class YamlReader() {
  private val mapper = new ObjectMapper(new YAMLFactory()).registerModule(DefaultScalaModule)
  mapper.registerModule(DefaultScalaModule)
  private val parseTypeRef = new TypeReference[Map[String, KeyObjectBuilder]]() {}

  def parse(file: File): List[KeyObject] = {
    val dataSource = Source.fromFile(file)
    val yamlData = dataSource.mkString
    dataSource.close()
    parseFromContent(yamlData)
  }

  def parse(resourcePath: String): List[KeyObject] = {
    val yamlData = Source.fromResource(resourcePath).mkString
    parseFromContent(yamlData)
  }

  private def parseFromContent(content: String) = {
    val yamlMap: Map[String, KeyObjectBuilder] = mapper.readValue(content, parseTypeRef)
    yamlMap.map { case (yamlKey, keyObject) =>
      keyObject.key = yamlKey
      keyObject.build()
    }.toList
  }
}
