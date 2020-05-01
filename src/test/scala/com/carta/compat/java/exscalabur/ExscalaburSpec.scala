package com.carta.compat.java.exscalabur

import java.io.ByteArrayOutputStream

import com.carta.mock.MockExscalabur
import org.scalatest.BeforeAndAfterEach
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

class ExscalaburSpec extends AnyFlatSpec with Matchers with BeforeAndAfterEach {
  var mockScalaExscalabur: MockExscalabur = new MockExscalabur()
  var javaExscalabur: Exscalabur = new Exscalabur(mockScalaExscalabur)

  "Exscalabur.getAppendOnlySheetWriter" should "call scala exscalabur's getAppendOnlySheetWriter method" in {
    javaExscalabur.getAppendOnlySheetWriter("some sheet name")
    mockScalaExscalabur.getAppendOnlySheetWriterCallCount shouldBe 1
  }

  "Exscalabur.exportToFile" should "call scala exscalabur's exportToFile method" in {
    javaExscalabur.exportToFile("some file path")
    mockScalaExscalabur.exportToFileCallCount shouldBe 1
  }

  "Exscalabur.exportToStream" should "call scala exscalabur's exportToStream method" in {
    javaExscalabur.exportToStream(new ByteArrayOutputStream())
    mockScalaExscalabur.exportToStreamCallCount shouldBe 1
  }

  override protected def beforeEach(): Unit = {
    super.beforeEach()
    mockScalaExscalabur.reset()
  }
}
