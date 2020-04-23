[![Scala version](https://img.shields.io/badge/scala-2.12.8-brightgreen.svg)](https://www.scala-lang.org/download/2.12.8.html)

# Exscalabur
Exscalabur is a Scala library for creating Excel files from data and a template. Exscalabur allows excel layout and formatting to be specified by excel templates, allowing anyone with basic excel knowledge to design and modify templates.

## Development

### Requirements

- Java 11 or later
- Scala 2.12 or later
- [sbt](https://www.scala-sbt.org/1.x/docs/Setup.html) 1.2.8 or higher

### Common sbt Commands

| Command     | Description                                          |
| ----------- | ---------------------------------------------------- |
| `~compile`  | `~` enables hot reloading                            |
| `~run`      | `~` enables hot reloading                            |
| `test`      | Runs all tests                                       |
| `testQuick` | Runs tests only affected by your latest code changes |
| `clean`     | Removes generated files from the target directory    |
| `update`    | Updates external dependencies                        |
| `package`   | Creates JAR                                          |
| `publishLocal` | Publishes to local IVY repository, by default in `~/.ivy2/local` |

## How To Use

## Templates
To use Exscalabur to create Excel files, you first require an __Excel Template__ file. 

The Excel template contains formatting and layout but instead of data, there are __keys__ in the cells that indicate where the data will be injected.

Any cell styling/cell borders made to the template will appear in the generated excel file.

An example template sheet:

![template](.readme_resources/template.png)

## Template Keys
A template key is a cell with the value prefixed by `$KEY.` or `$REP.`
* A `$KEY.` cell will be substituted for a single piece of data.
* A `$REP.` cell will be repeated with formatting for each row of data provided.

Cells that do not contain keys will be copied as is.

In the above example, this template has 6 keys:`$KEY.fname`, `$KEY.lname`, `$REP.animal`, `$REP.weight`, `$KEY.conclusion`,and `$REP.element`.

## Schema Definition
### Yaml
A schema definition is required, and may be provided as a YAML file with the structure:

```yaml
KEYNAME1:
  dataType: oneOf("string", "double", "long", "date")
  excelType: oneOf("string", "number", "date")
KEYNAME2:
  dataType: oneOf("string", "double", "long")
  excelType: oneOf("string", "number", "date")  
# etc
```

* `KEYNAME` is the **key** exactly as seen in the template.
* `dataType` is the runtime type of the data provided to exscalabur. 
* `excelType` is the type of the cell in the final sheet.

`com.carta.yaml.YamlReader` is provided to read the schema definition yaml files. 

For the example template Excel file above, the following is valid `schema` definition:

```yaml
$KEY.fname:
  columnType: "string"
  excelType: "string"

$KEY.lname:
  columnType: "string"
  excelType: "string"

$KEY.conclusion:
  columnType: "string"
  excelType: "string"

$REP.animal:
  columnType: "string"
  excelType: "string"

$REP.weight:
  columnType: "double"
  excelType: "double"
  
$REP.element:
  columnType: "string"
  excelType: "string"
```

### In-Code Map
Alternatively, this **schema** may be provided in-code as a `Map[String, YamlEntry]`.
The Map's keys are the `KEYNAME`s.
`import com.carta.yaml.YamlEntry` is a provided case class representing the above YAML structure.  

## Data Substitution

### Single-Substitution Data

To pass data in to be substituted into a `$KEY.` template cell, an instance of a `DataCell(key: String, value: oneOf(String, Long, Double))` must be created.
`key` is the __key__ from the template, __without the `$KEY.` prefix__.
`value` is the value to substitute. It's runtime type must match the schema definition.

For the above example, we may have `DataCell`s like:

```Scala
import com.carta.exscalabur.DataCell

DataCell("fname", "Joe");
DataCell("lname", "Person")
DataCell("conclusion", "EXSCALABUR")
```

### Repeated Data Substitution
To pass data to be substituted into a repeated row, instances of `DataRow` are passed in, containing `DataCell`s for each cell in the repeated row.

For the example, our `DataRow` instances might look like:

```scala
import com.carta.exscalabur.DataRow

DataRow().addCell("animal", "monkey").addCell("weight", 12.1)
DataRow().addCell("animal", "horse").addCell("weight", 12.2)
DataRow().addCell("element", "hydrogen")
DataRow().addCell("element", "helium")
DataRow().addCell("element", "lithium")
```

## Writing

Exscalabur supports multi-step writes. This can be done with multiple calls to `writeStaticData` and `writeRepeatedData`.

Currently, Exscalabur only supports writing in an append-only manner. So, for the above example, the data for `$KEY.fname` must be provided before the data for $REP.animal is provided.

There are plans for Exscalabur to support writing to rows out of order, but this has yet to be implemented.

__Exscalabur does not support sub-tables arranged horizontally__

Lastly, all that's left is to write the data. To do so, create an instance of a `Exscalabur` object:

```scala
import com.carta.exscalabur.Exscalabur

Exscalabur(
  templates, // Iterable[String], representing paths to the template sheets
  schema, // The in-code schema representation explained above
  windowSize // number of rows in the output workbook to keep in memory at a time
)
```

For every sheet to be written to, create a `SheetWriter`:

```scala
exscalabur.getAppendOnlySheetWriter(sheetName) // sheetName is as defined in the template sheet.
```
Note, the `sheetName` used here comes from the tab in the template. It __must match the template exactly__

Write static data: `sheetWriter.writeStaticData(staticData: Seq[DataCell])`

and to write repeated data: `sheetWriter.writeRepeatedData(repeatedData: Seq[DataRow])`

And write to the output file: `exscalabur.exportToFile(outputPath)`

Doing so results in the final output sheet:

![output](.readme_resources/output.png)

# Supported Cell Types
Currently, Exscalabur supports copying cells containing numeric, boolean, and string values. 
Cells containing formulas and formula errors will not copied from the template to the output sheet.

# Roadmap
* Add support for Exscalabur to copy formulas on single-substitution template cells.