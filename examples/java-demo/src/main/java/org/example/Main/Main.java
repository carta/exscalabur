package org.example.Main;

import com.carta.compat.java.excel.AppendOnlySheetWriter;
import com.carta.compat.java.exscalabur.DataCell;
import com.carta.compat.java.exscalabur.DataRow;
import com.carta.compat.java.exscalabur.Exscalabur;
import com.carta.compat.java.yaml.YamlReader;
import com.carta.yaml.YamlEntry;
import org.apache.commons.compress.utils.Lists;

import java.util.ArrayList;
import java.util.Arrays;
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

        final List<String> names = Arrays.asList("Jon Somebody", "Jane Person", "Frank LastName", "Sarah Fakename", "Tim Tom");
        final List<Double> gpa = Arrays.asList(3.8, 4.2, 4.1, 3.9, 3.9);
        final List<String> major = Arrays.asList("CS", "Engineering", "Arts", "Science", "Engineering");

        final List<DataRow> peopleData1 = Lists.newArrayList();
        final List<DataRow> peopleData2 = Lists.newArrayList();

        for (int i = 0; i < 3; i++) {
            DataRow row = new DataRow();
            row.addCell("person", names.get(i));
            row.addCell("gpa", gpa.get(i));
            row.addCell("major", major.get(i));
            peopleData1.add(row);
        }

        for (int i = 3; i < 5; i++) {
            DataRow row = new DataRow();
            row.addCell("person", names.get(i));
            row.addCell("gpa", gpa.get(i));
            row.addCell("major", major.get(i));
            peopleData2.add(row);
        }

        final List<String> schoolNames = Arrays.asList("School Name", "University of Place", "Well Known School");
        final List<Long> numStudents = Arrays.asList(43_000L, 25_000L, 30_000L);
        final List<DataRow> schoolData = Lists.newArrayList();
        for (int i = 0; i < schoolNames.size(); i++) {
            DataRow row = new DataRow();
            row.addCell("school", schoolNames.get(i));
            row.addCell("numStudents", numStudents.get(i));
            schoolData.add(row);
        }

        final List<DataCell> staticData = Lists.newArrayList();
        staticData.add(new DataCell("project", "exscalabur demo"));

        final List<DataCell> staticData2 = Lists.newArrayList();
        staticData2.add(new DataCell("value1", 12.34));
        staticData2.add(new DataCell("value2", 11.35));
        staticData2.add(new DataCell("value3", 14.50));

        final AppendOnlySheetWriter sheetWriter = sword.getAppendOnlySheetWriter("Sheet1");

        sheetWriter.writeStaticData(staticData);
        sheetWriter.writeRepeatedData(peopleData1);
        sheetWriter.writeRepeatedData(peopleData2);
        sheetWriter.writeRepeatedData(schoolData);
        sheetWriter.writeStaticData(staticData2);

        sword.exportToFile("./out.xlsx");
    }
}