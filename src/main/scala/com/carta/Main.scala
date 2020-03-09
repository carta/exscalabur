package com.carta

import com.carta.exscalabur.{DataRow, Exscalabur}

import scala.collection.immutable

object Main extends App {
  val sword = new Exscalabur(
    "/Users/jacksonlo/Downloads/exscalabur.xlsx",
    "/Users/jacksonlo/Dev/exscalabur/src/test/resources/test.yaml"
    )

  val row1 = DataRow.Builder()
    .addCell("string", "foo")
    .addCell("numbers", 12)
    .addCell("dates", 1581028273)
    .build()

  val single = DataRow.Builder()
      .addCell("company_name", "2019 Q1 HACKATHON")
      .addCell("coop_one", "Ziyad AlYafi")
      .addCell("coop_two", "Het Kataria")
    .build()

  val animals = List("bear", "eagle", "elephant", "bird", "snake", "pig", "dog", "cat", "penguin", "anteater");
  val numbers = List(1, 2, 3, 4, 5, 6, 7, 8, 9, 10);
  val foods = List("ice cream", "sushi", "curry", "sandwich", "burger", "candy", "carrots", "avocado", "la croix", "potatoes");
  val dates = List("2019/01/01", "2019/01/02", "2019/01/03", "2019/01/04", "2019/01/05", "2019/01/06", "2019/01/07", "2019/01/08", "2019/01/09", "2019/01/10")
  val cities = List("Toronto", "Waterloo", "Windsor", "San Francisco", "Palo Alto", "Washington", "Kitchener", "Rio de Janeiro", "SLC", "NYC")

  val rows: immutable.Seq[DataRow] = (0 until 10).map(i => {
    DataRow.Builder()
      .addCell("animals", animals(i))
      .addCell("numbers", numbers(i))
      .addCell("foods", foods(i))
      .addCell("dates", dates(i))
      .addCell("cities", cities(i))
      .build()
  });

  sword.addTab(
    "lots of formatting",
    "/Users/jacksonlo/Downloads/demo_template.xlsx",
    single,
    rows.toList
  )

  sword.writeExcelToDisk()
}
