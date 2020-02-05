package com.carta.excel

import com.carta.exscalabur.DataCell
import com.carta.yaml.{DataType, ExcelType, YamlEntry}
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

class ModelMapSpec extends AnyFlatSpec with Matchers {
  "ModelMap" should "produce a blank cell if undefined key is provided" in {
    val keyType = "REP"
    val dataRow = List(DataCell("undefined", 1))
    val schema = Map("$REP.defined" -> YamlEntry(DataType.string, ExcelType.string))

    val modelMap = ModelMap(keyType, dataRow, schema)
    modelMap.values.foreach(cellValue => assert(cellValue.isInstanceOf[CellBlank]))
  }

  "ModelMap" should "produce a blank cell if schema is empty" in {
    val keyType = "REP"
    val dataRow = List(DataCell("undefined", 1))
    val schema = Map.empty[String, YamlEntry]

    val modelMap = ModelMap(keyType, dataRow, schema)
    modelMap.values.foreach(cellValue => assert(cellValue.isInstanceOf[CellBlank]))
  }

  List(("test", DataType.string), (1L, DataType.long), (1.5, DataType.double)).foreach { case (value, dataType) =>
    "ModelMap" should f"produce a string cell excel type is string for dataType $dataType" in {
      val keyType = "$REP"
      val dataRow = value match {
        case s: String => List(DataCell("defined", s))
        case l: Long => List(DataCell("defined", l))
        case d: Double => List(DataCell("defined", d))
      }
      val schema = Map("$REP.defined" -> YamlEntry(dataType, ExcelType.string))

      val modelMap = ModelMap(keyType, dataRow, schema)
      modelMap.values.foreach(cellValue => assert(cellValue.isInstanceOf[CellString]))
    }
  }
}
