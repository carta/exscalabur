import com.carta.exscalabur.{DataCell, DataRow, Exscalabur}
import com.carta.yaml.YamlReader

object Main {
  def main(args: Array[String]): Unit = {
    val templates: List[String] = List(getClass.getResource("/excel/template.xlsx").getPath)

    val schemaPath = getClass.getResource("/yaml/schema.yaml").getPath
    val schemaDefinition = YamlReader().parse(schemaPath)

    val sword = Exscalabur(
      templates,
      schemaDefinition,
      windowSize = 100
    )

    val peopleData1: List[DataRow] = List(
      DataRow().addCell("person", "Jon Somebody").addCell("gpa", 3.8).addCell("major", "CS"),
      DataRow().addCell("person", "Jane Person").addCell("gpa", 4.2).addCell("major", "Engineering"),
      DataRow().addCell("person", "Frank LastName").addCell("gpa", 4.1).addCell("major", "Arts"),
    )

    val peopleData2: List[DataRow] = List(
      DataRow().addCell("person", "Sarah Fakename").addCell("gpa", 3.9).addCell("major", "Science"),
      DataRow().addCell("person", "Tim Tom").addCell("gpa", 3.9).addCell("major", "Engineering"),
    )

    val schoolData: List[DataRow] = List(
      DataRow().addCell("school", "School Name").addCell("numStudents", 43000),
      DataRow().addCell("school", "University of Place").addCell("numStudents", 25000),
      DataRow().addCell("school", "Well Known School").addCell("numStudents", 30000)
    )

    val staticData = List(DataCell("project", "exscalabur demo"))
    val sheetWriter = sword.getAppendOnlySheetWriter("Sheet1")

    sheetWriter.writeStaticData(staticData)
    sheetWriter.writeRepeatedData(peopleData1)
    sheetWriter.writeRepeatedData(peopleData2)
    sheetWriter.writeRepeatedData(schoolData)

    sword.exportToFile("./out.xlsx")
  }
}