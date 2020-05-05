package com.carta.mock

import com.carta.excel.CellFormulaParser
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import scala.collection.mutable

class MockAppendOnlySheetWriter extends com.carta.excel.AppendOnlySheetWriter(
  templateSheet = new XSSFWorkbook().createSheet(),
  outputWorkbook = new XSSFWorkbook(),
  schema = Map.empty,
  cellStyleCache = mutable.Map.empty,
  cellFormulaParser = new CellFormulaParser
) {
  var writeStaticDataCallCount = 0
  var writeRepeatedDataCallCount = 0
  var copyPicturesCallCount = 0

  def reset(): Unit = {
    writeStaticDataCallCount = 0
    writeRepeatedDataCallCount = 0
    copyPicturesCallCount = 0
  }

  override def writeStaticData(data: Seq[com.carta.exscalabur.DataCell]): Unit = writeStaticDataCallCount += 1

  override def copyPictures(): Unit = copyPicturesCallCount += 1

  override def writeRepeatedData(data: Seq[com.carta.exscalabur.DataRow]): Unit = writeRepeatedDataCallCount += 1
}