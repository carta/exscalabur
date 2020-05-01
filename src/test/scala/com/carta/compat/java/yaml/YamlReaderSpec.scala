package com.carta.compat.java.yaml

import java.io.File

import com.carta.mock.MockYamlReader
import org.scalatest.BeforeAndAfterEach
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

class YamlReaderSpec extends AnyFlatSpec with Matchers with BeforeAndAfterEach {
  val mockYamlReader = new MockYamlReader()
  val yamlReader = new YamlReader(mockYamlReader)

  "YamlReader.parse(file)" should "call scala YamlReader's parse method" in {
    yamlReader.parse(new File("."))
    mockYamlReader.parseFileCallCount shouldBe 1
  }

  "YamlReader.parse(path)" should "call scala YamlReader's parse method" in {
    yamlReader.parse("some path")
    mockYamlReader.parsePathCallCount shouldBe 1
  }

  override protected def beforeEach(): Unit = {
    mockYamlReader.reset()
  }
}
