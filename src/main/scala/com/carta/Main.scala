package com.carta

import com.carta.temp.{DataRow, Exscalabur}

object Main extends App {
  val sword = new Exscalabur(
    "/tmp/exscalabur.xlsx",
    "/Users/daviddufour/code/hackathon/exscalabur/src/test/resources/test.yaml"
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

  val row3 = DataRow.Builder()
    .addCell("string", "hello")
    .addCell("numbers", 1337)
    .addCell("dates", 1520028273)
    .build()

  sword.addTab(
    "tab1_name",
    "/Users/daviddufour/code/hackathon/exscalabur/src/resources/templates/tab1_template.xlsx",
    DataRow.Builder()
           .addCell("number", 12)
           .build(),
    List()
    )
  sword.addTab(
    "tab2_name",
    "/Users/daviddufour/code/hackathon/exscalabur/src/resources/templates/tab2_template.xlsx",
    DataRow.Builder()
      .addCell("number", 24)
      .build(),
    List(row1, row2, row3)
    )

  sword.writeExcelToDisk()
}
