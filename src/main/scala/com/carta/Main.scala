package com.carta

import com.carta.temp.{DataRow, Exscalabur}

object Main extends App {
  val sword = new Exscalabur(
    "/Users/katiedsouza/Desktop/exscalabur.xlsx",
    "/Users/katiedsouza/Developer/exscalabur/src/test/resources/test.yaml"
  )

  val row1 = DataRow.Builder()
    .addCell("string", "foo")
    .addCell("numbers", 12)
    .addCell("dates", 1581028273)
    .build()

  val row2 = DataRow.Builder()
    .addCell("string", "bar")
    .addCell("numbers", 13)
    .addCell("dates", 1580028273)
    .build()

  sword.addTab(
    "tab2_name",
    "/Users/katiedsouza/Developer/exscalabur/src/resources/templates/tab2_template.xlsx",
    List(row1, row2)
  ).writeExcelToDisk()
}
