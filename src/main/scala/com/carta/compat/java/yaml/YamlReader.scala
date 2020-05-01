package com.carta.compat.java.yaml

import java.io.File
import scala.collection.JavaConverters._
import com.carta.yaml.YamlEntry

class YamlReader(scalaYamlReader: com.carta.yaml.YamlReader) {
  def this() {
    this(com.carta.yaml.YamlReader())
  }

  def parse(file: File): java.util.Map[String, YamlEntry] = {
    scalaYamlReader.parse(file).asJava
  }

  def parse(path: String): java.util.Map[String, YamlEntry] = {
    scalaYamlReader.parse(path).asJava
  }
}
