package com.carta.yaml

import java.io.File

import com.carta.UnitSpec

import scala.collection.mutable.Stack

class YamlReaderSpec extends UnitSpec {
  val testDataPath = "test.yaml"
  val testDataFile = new File(getClass.getClassLoader.getResource(testDataPath).toURI)

  "YAML Reader" should "Produce YamlEntries given yaml file" in {
    val expectedData = Vector(
      YamlEntry(KeyType.single, CellType.string, CellType.string),
      YamlEntry(KeyType.repeated, CellType.double, CellType.double),
      YamlEntry(KeyType.single, CellType.long, CellType.date)
    )

    val yamlReader = YamlReader()
    val actualData = yamlReader.parse(testDataPath).values

    actualData should contain theSameElementsAs expectedData
  }

  it should "Produce KeyObject from yaml file as File" in {
    val keyObjects = YamlReader().parse(testDataFile)

    val expectedData = Vector(
      YamlEntry(KeyType.single, CellType.string, CellType.string),
      YamlEntry(KeyType.repeated, CellType.double, CellType.double),
      YamlEntry(KeyType.single, CellType.long, CellType.date)
    )

    keyObjects.values should contain theSameElementsAs expectedData
  }
}