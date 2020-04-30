package org.example.Main;

import com.carta.compat.java.excel.AppendOnlySheetWriter;
import com.carta.compat.java.exscalabur.DataCell;
import com.carta.compat.java.exscalabur.DataRow;
import com.carta.compat.java.exscalabur.Exscalabur;
import com.carta.compat.java.yaml.YamlReader;
import com.carta.yaml.YamlEntry;
import org.apache.commons.compress.utils.Lists;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

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

        final Map<String, YamlEntry> schemaDefinition = new YamlReader().parse(schemaPath);

        final Exscalabur sword = new Exscalabur(templates, schemaDefinition, 100);

        final List<DataRow> peopleData1 = Lists.newArrayList();
        peopleData1.add(new DataRow().addCell("person", "Jon Somebody").addCell("gpa", 3.8).addCell("major", "CS"));
        peopleData1.add(new DataRow().addCell("person", "Jane Person").addCell("gpa", 4.2).addCell("major", "Engineering"));
        peopleData1.add(new DataRow().addCell("person", "Frank LastName").addCell("gpa", 4.1).addCell("major", "Arts"));

        final List<DataRow> peopleData2 = Lists.newArrayList();
        peopleData2.add(new DataRow().addCell("person", "Sarah Fakename").addCell("gpa", 3.9).addCell("major", "Science"));
        peopleData2.add(new DataRow().addCell("person", "Tim Tom").addCell("gpa", 3.9).addCell("major", "Engineering"));

        final List<DataRow> schoolData = Lists.newArrayList();
        schoolData.add(new DataRow().addCell("school", "School Name").addCell("numStudents", 43000));
        schoolData.add(new DataRow().addCell("school", "University of Place").addCell("numStudents", 25000));
        schoolData.add(new DataRow().addCell("school", "Well Known School").addCell("numStudents", 30000));

        final List<DataCell> staticData = Lists.newArrayList();
        staticData.add(new DataCell("project", "exscalabur demo"));


        final AppendOnlySheetWriter sheetWriter = sword.getAppendOnlySheetWriter("Sheet1");

        sheetWriter.writeStaticData(staticData);
        sheetWriter.writeRepeatedData(peopleData1);
        sheetWriter.writeRepeatedData(peopleData2);
        sheetWriter.writeRepeatedData(schoolData);

        sword.exportToFile("./out.xlsx");
    }
}