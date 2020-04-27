package org.example.Main;

import com.carta.excel.AppendOnlySheetWriter;
import com.carta.exscalabur.DataCell;
import com.carta.exscalabur.DataRow;
import com.carta.exscalabur.Exscalabur;
import com.carta.yaml.YamlEntry;
import com.carta.yaml.YamlReader;
import org.apache.commons.compress.utils.Lists;
import scala.collection.JavaConverters;
import scala.collection.immutable.Map;

import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        new App().run();
    }
}

class App {
    public void run() {
        final ClassLoader classLoader = this.getClass().getClassLoader();
        final String templatePath = classLoader.getResource("excel/template.xlsx").getFile();
        final String schemaPath = classLoader.getResource("yaml/schema.yaml").getFile();

        final ArrayList<String> templates = Lists.newArrayList();
        templates.add(templatePath);

        final Map<String, YamlEntry> schemaDefinition = YamlReader.apply().parse(schemaPath);

        final Exscalabur sword = Exscalabur.apply(
                JavaConverters.collectionAsScalaIterable(templates),
                schemaDefinition,
                100
        );

        final List<DataRow> peopleData1 = Lists.newArrayList();
        peopleData1.add(DataRow.apply().addCell("person", "Jon Somebody").addCell("gpa", 3.8).addCell("major", "CS"));
        peopleData1.add(DataRow.apply().addCell("person", "Jane Person").addCell("gpa", 4.2).addCell("major", "Engineering"));
        peopleData1.add(DataRow.apply().addCell("person", "Frank LastName").addCell("gpa", 4.1).addCell("major", "Arts"));

        final List<DataRow> peopleData2 = Lists.newArrayList();
        peopleData2.add(DataRow.apply().addCell("person", "Sarah Fakename").addCell("gpa", 3.9).addCell("major", "Science"));
        peopleData2.add(DataRow.apply().addCell("person", "Tim Tom").addCell("gpa", 3.9).addCell("major", "Engineering"));

        final List<DataRow> schoolData = Lists.newArrayList();
        schoolData.add(DataRow.apply().addCell("school", "School Name").addCell("numStudents", 43000));
        schoolData.add(DataRow.apply().addCell("school", "University of Place").addCell("numStudents", 25000));
        schoolData.add(DataRow.apply().addCell("school", "Well Known School").addCell("numStudents", 30000));

        final List<DataCell> staticData = Lists.newArrayList();
        staticData.add(DataCell.apply("project", "exscalabur demo"));


        final AppendOnlySheetWriter sheetWriter = sword.getAppendOnlySheetWriter("Sheet1");

        sheetWriter.writeStaticData(JavaConverters.iterableAsScalaIterable(staticData).toSeq());
        sheetWriter.writeRepeatedData(JavaConverters.iterableAsScalaIterable(peopleData1).toSeq());
        sheetWriter.writeRepeatedData(JavaConverters.iterableAsScalaIterable(peopleData2).toSeq());
        sheetWriter.writeRepeatedData(JavaConverters.iterableAsScalaIterable(schoolData).toSeq());

        sword.exportToFile("./out.xlsx");
    }
}