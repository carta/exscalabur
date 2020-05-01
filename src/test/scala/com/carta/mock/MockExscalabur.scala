package com.carta.mock

import java.io.OutputStream

import com.carta.excel.AppendOnlySheetWriter
import org.apache.poi.xssf.streaming.SXSSFWorkbook

import scala.collection.mutable

class MockExscalabur extends com.carta.exscalabur.Exscalabur(
  Map.empty,
  new SXSSFWorkbook(100),
  mutable.Map.empty,
  Map.empty
) {
  var getAppendOnlySheetWriterCallCount = 0
  var exportToFileCallCount = 0
  var exportToStreamCallCount = 0

  def reset(): Unit = {
    getAppendOnlySheetWriterCallCount = 0
    exportToFileCallCount = 0
    exportToStreamCallCount = 0
  }

  override def getAppendOnlySheetWriter(sheetName: String): AppendOnlySheetWriter = {
    getAppendOnlySheetWriterCallCount += 1
    new MockAppendOnlySheetWriter()
  }

  override def exportToFile(path: String): Unit = exportToFileCallCount+=1

  override def exportToStream(outputStream: OutputStream): Unit = exportToStreamCallCount+=1
}