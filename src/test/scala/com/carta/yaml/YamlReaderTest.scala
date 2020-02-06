package com.carta.yaml

import java.io.File

import com.carta.UnitSpec

import scala.collection.mutable.Stack

class YamlReaderSpec extends UnitSpec {

  "YAML Reader" should "Produce KeyObject from yaml file as resource" in {
    val yamlReader = new YamlReader()
    val keyObjects = yamlReader.parse("test.yaml")

    val expectedKey1Obj = KeyObject(
      key = "key1",
      keyType = KeyType.single,
      columnType = CellType.double,
      excelType = CellType.double
    )

    val expectedKey2Obj = KeyObject(
      key = "key2",
      keyType = KeyType.repeated,
      columnType = CellType.long,
      excelType = CellType.date
    )

    keyObjects.length shouldBe 2

    var actualKey1Obj: KeyObject = null
    var actualKey2Obj: KeyObject = null

    if (keyObjects.head.key == "key1") {
      actualKey1Obj = keyObjects(0)
      actualKey2Obj = keyObjects(1)
    }
    else {
      actualKey1Obj = keyObjects(1)
      actualKey2Obj = keyObjects(0)
    }
    actualKey1Obj shouldBe expectedKey1Obj
    actualKey2Obj shouldBe expectedKey2Obj
  }

  it should "Produce KeyObject from yaml file as File" in {
    val yamlReader = new YamlReader()
    val yamlFile = new File(getClass.getClassLoader.getResource("test.yaml").toURI)
    val keyObjects = yamlReader.parse(yamlFile)

    val expectedKey1Obj = KeyObject(
      key = "key1",
      keyType = KeyType.single,
      columnType = CellType.double,
      excelType = CellType.double
    )

    val expectedKey2Obj = KeyObject(
      key = "key2",
      keyType = KeyType.repeated,
      columnType = CellType.long,
      excelType = CellType.date
    )

    keyObjects.length shouldBe 2

    var actualKey1Obj: KeyObject = null
    var actualKey2Obj: KeyObject = null

    if (keyObjects.head.key == "key1") {
      actualKey1Obj = keyObjects(0)
      actualKey2Obj = keyObjects(1)
    }
    else {
      actualKey1Obj = keyObjects(1)
      actualKey2Obj = keyObjects(0)
    }
    actualKey1Obj shouldBe expectedKey1Obj
    actualKey2Obj shouldBe expectedKey2Obj
  }
}