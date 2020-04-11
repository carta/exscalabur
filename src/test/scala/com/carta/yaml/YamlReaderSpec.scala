package com.carta.yaml

import java.io.File

import com.carta.UnitSpec

class YamlReaderSpec extends UnitSpec {
  val testDataPath: String = getClass.getResource("/test.yaml").getFile
  val testDataFile: File = new File(testDataPath)

  "YAML Reader" should "Produce YamlEntries given yaml file" in {
    val expectedData = Vector(
      YamlEntry(KeyType.single, YamlCellType.string, YamlCellType.string),
      YamlEntry(KeyType.repeated, YamlCellType.double, YamlCellType.double),
      YamlEntry(KeyType.single, YamlCellType.long, YamlCellType.date)
    )

    val yamlReader = YamlReader()
    val actualData = yamlReader.parse(testDataPath).values

    actualData should contain theSameElementsAs expectedData
  }

  it should "Produce KeyObject from yaml file as File" in {
    val keyObjects = YamlReader().parse(testDataFile)

    val expectedData = Vector(
      YamlEntry(KeyType.single, YamlCellType.string, YamlCellType.string),
      YamlEntry(KeyType.repeated, YamlCellType.double, YamlCellType.double),
      YamlEntry(KeyType.single, YamlCellType.long, YamlCellType.date)
    )

    keyObjects.values should contain theSameElementsAs expectedData
  }
}