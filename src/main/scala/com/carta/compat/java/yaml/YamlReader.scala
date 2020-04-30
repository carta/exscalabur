package com.carta.compat.java.yaml

import java.io.File
import scala.collection.JavaConverters._
import com.carta.yaml.{YamlEntry, YamlReader}

class YamlReader {
  val scalaYamlReader = YamlReader()

  def parse(file: File): java.util.Map[String, YamlEntry] = {
    scalaYamlReader.parse(file).asJava
  }

  def parse(path: String): java.util.Map[String, YamlEntry] = {
    scalaYamlReader.parse(path).asJava
  }
}
