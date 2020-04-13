[![CircleCI](https://circleci.com/gh/carta/exscalabur/tree/master.svg?style=svg&circle-token=a8e8f68d2e70a177a3298140e5ec935710f651c7)](https://circleci.com/gh/carta/exscalabur/tree/master)
[![Scala version](https://img.shields.io/badge/scala-2.12.8-brightgreen.svg)](https://www.scala-lang.org/download/2.12.8.html)
[![Latest Version @ Cloudsmith](https://api-prd.cloudsmith.io/badges/version/carta/maven-releases/maven/exscalabur_2.12/latest/xg=com.carta/?render=true&badge_token=gAAAAABePJwL9tZMOa6DrXC6N_iGkYROA2I1jTSwarIRvAuhy7O34Tt742-doost6rUHEs5WR2PqRoxjGCihc1v0mCeHIeVY_hSi6-wyPttrjUAFaGPmXMU%3D)](https://cloudsmith.io/~carta/repos/maven-releases/packages/detail/maven/exscalabur_2.12/latest/xg=com.carta/)

# Exscalabur

A Scala library for creating excel files from data and a template.

## Usage

Add it as a dependency:
`libraryDependencies += "com.carta" % "exscalabur" % "2.12.8"`

## Development

### Requirements

- Java 12 or later
- Scala 2.12.8
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

| `publishLocal` | Publishes to local IVY repository, by default in `~/.ivy2/local`

## How To Use

To use Exscalabur to create Excel files, you first require an __Excel Template__ file. This file is a bare-bones version of what the final Excel file will look like, but with special __keys__ specifying where data will be located.

An example of what a template sheet may look like:

![template](/Users/katiedsouza/Developer/exscalabur/.readme_resources/template.png)

This template has 6 **keys** for data substitution --`$KEY.fname`, `$KEY.lname`, `$REP.animal`, `$REP.weight`, `$KEY.conclusion`,and `$REP.element`.  A **key** is defined by having either a `$KEY.`, or `$REP.` prefix. If a row contains a cell with a `$REP.*` cell, it is considered a __repeated row__ -- one that will be repeated for every piece of data supplied. If a cell contains a `$KEY.`, a one-time substitution will be applied for data that is provided. Anything not a __key__ will be copied as is. 

Any cell styling/cell borders must be made to the template.

The next thing required is a __schema definition__. This may be provided as a yaml file. The format of the Yaml file is as follows:

```yaml
KEYNAME1:
  dataType: oneOf("string", "double", "long")
  excelType: oneOf("string", "number", "date")
KEYNAME2:
  dataType: oneOf("string", "double", "long")
  excelType: oneOf("string", "number", "date")  
# etc
```

Where `KEYNAME` is a **key** as seen in the template sheet, `excelType` is the type of the cell in the final sheet, and `dataType`, is the type of the data that will be passed in.

A  `YamlReader` class is provided to convert the yaml file into an in-code representation. 

Alternatively, this **schema** may be provided in-code as a `Map[String, YamlEntry]`, where `YamlEntry` is a provided case class representing the above structure, and the map's keys represent the `KEYNAME`s of the above example.

For the above template Excel file, the following may be an example of the `schema` definition:

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

To pass data in to be substituted into a `$KEY.` template cell, an instance of a `DataCell(key: String, value: oneOf(String, Long, Double))` must be created, where the key represents a __key__ _without the `$KEY.` prefix_, and `value` is the value to substitute.

For example, in the above example, we may have `DataCell`s like:

```Scala
DataCell("fname", "Joe");
DataCell("lname", "Person")
DataCell("conclusion", "EXSCALABUR")
```

To pass data in to be substituted into a repeated row, instances of `DataRow` are passed in, containing `DataCell`s for each cell in the repeated row. A `Builder` is provided to construct these instances.

For the example, our `DataRow` instances might look like:

```scala
DataRow.Builder().addCell("animal", "monkey").addCell("weight", 12.1).build()
DataRow.Builder().addCell("animal", "horse").addCell("weight", 12.2).build()
DataRow.Builder().addCell("element", "hydrogen").build()
DataRow.Builder().addCell("element", "helium").build()
DataRow.Builder().addCell("element", "lithium").build()
```

Exscalabur supports performing writes in multiple steps. So, data that is to be written is supplied, wrapped in a `Iterator[(List[DataCell], List[DataRow])]`. As such, the data for `$REP.animal` may be provided in multiple (sequential) elements of this `Iterator`.

Currently, Exscalabur only supports writing in an append-only manner. So, for the above example, the data for `$KEY.fname` cannot be provided any later than the first set of data provided  for `$REP.animal`.

There are plans for Exscalabur to support writing to rows out of order, but this has yet to be implemented.

Lastly, all that's left is to write the data. To do so, create an instance of a `Exscalabur` object:

```scala
Exscalabur(
  templates, // Iterable[String], representing paths to the template sheets
  "/home/my_user/out.xlsx", // output file path
  yamlData, // The in-code schema representation explained above
  windowSize // number of rows in the output workbook to keep in memory at a time
)
```

For every sheet to be written to, create a `SheetWriter`: 

```scala
exscalabur.getAppendOnlySheetWriter(sheetName) // sheetName is as defined in the template sheet.
```

Write data: `sheetWriter.writeData(dataProvider: Iterable[(List[DataCell], List[DataRow])])`

And write to the output file: `exscalabur.writeToDisk()`

Doing so results in the final output sheet:

![output](/Users/katiedsouza/Developer/exscalabur/.readme_resources/output.png)