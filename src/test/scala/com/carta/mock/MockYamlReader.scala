package com.carta.mock

import java.io.File

import com.carta.yaml.YamlEntry
import com.fasterxml.jackson.databind.ObjectMapper

class MockYamlReader extends com.carta.yaml.YamlReader(new ObjectMapper()) {
  var parseFileCallCount = 0
  var parsePathCallCount = 0

  def reset(): Unit = {
    parseFileCallCount = 0
    parsePathCallCount = 0
  }

  override def parse(file: File): Map[String, YamlEntry] = {
    parseFileCallCount += 1
    Map.empty
  }

  override def parse(path: String): Map[String, YamlEntry] = {
    parsePathCallCount += 1
    Map.empty
  }
}
