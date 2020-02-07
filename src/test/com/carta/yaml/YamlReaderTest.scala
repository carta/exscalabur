package com.carta.yaml

import java.io.File

import com.carta.UnitSpec

import scala.collection.mutable.Stack

class YamlReaderSpec extends UnitSpec {

  "YAML Reader" should "Produce KeyObject from yaml file as resource" in {
    val yamlReader = new YamlReader()
    val keyObjects = yamlReader.parse("test.yaml")

    val expectedData = Seq(
      KeyObject(KeyType.repeated, CellType.long, CellType.double),
      KeyObject(KeyType.repeated, CellType.string, CellType.string),
      KeyObject(KeyType.repeated, CellType.long, CellType.double)
    )

    keyObjects.values.toSeq shouldEqual expectedData
  }

  it should "Produce KeyObject from yaml file as File" in {
    val yamlReader = new YamlReader()
    val yamlFile = new File(getClass.getClassLoader.getResource("test.yaml").toURI)
    val keyObjects = yamlReader.parse(yamlFile)

    val expectedData = Seq(
      KeyObject(KeyType.repeated, CellType.long, CellType.double),
      KeyObject(KeyType.repeated, CellType.string, CellType.string),
      KeyObject(KeyType.repeated, CellType.long, CellType.double)
    )

    keyObjects.values.toSeq shouldEqual expectedData
  }
}