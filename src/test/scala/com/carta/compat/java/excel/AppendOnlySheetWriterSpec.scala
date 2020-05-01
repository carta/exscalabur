package com.carta.compat.java.excel

import com.carta.compat.java.exscalabur.{DataCell, DataRow}
import com.carta.mock.MockAppendOnlySheetWriter
import org.scalatest.BeforeAndAfterEach
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

import scala.collection.JavaConverters._

class AppendOnlySheetWriterSpec extends AnyFlatSpec with Matchers with BeforeAndAfterEach {
  val mockScalaSheetWriter = new MockAppendOnlySheetWriter()
  val javaSheetWriter = new AppendOnlySheetWriter(mockScalaSheetWriter)

  "SheetWriter.writeStaticData" should "call scala sheet writer's writeStaticData method" in {
    val staticData = List(new DataCell("key", "value")).asJava
    javaSheetWriter.writeStaticData(staticData)
    mockScalaSheetWriter.writeStaticDataCallCount shouldBe 1
  }

  "SheetWriter.writeRepeatedData" should "call scala sheet writer's writeRepeatedData method" in {
    val repeatedData = List(new DataRow().addCell("key", "value")).asJava
    javaSheetWriter.writeRepeatedData(repeatedData)
    mockScalaSheetWriter.writeRepeatedDataCallCount shouldBe 1
  }

  "SheetWriter.copyPictures" should "call scala sheet writer's copyPictures method" in {
    javaSheetWriter.copyPictures()
    mockScalaSheetWriter.copyPicturesCallCount shouldBe 1
  }

  override protected def beforeEach(): Unit = {
    super.beforeEach()
    mockScalaSheetWriter.reset()
  }
}
