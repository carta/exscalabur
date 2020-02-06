import sbt.Keys.libraryDependencies

name := "Exscalabur"
organization := "carta"

version := "0.1"

scalaVersion := "2.12.8"

libraryDependencies ++= Seq(
  "org.scalatest" %% "scalatest" % "3.1.0" % "test",
  "org.scalactic" %% "scalactic" % "3.1.0",
  "org.apache.poi" % "poi" % "4.1.0",
  "org.apache.poi" % "poi-ooxml" % "4.1.0",
  "com.fasterxml.jackson.core" % "jackson-databind" % "2.10.2",
  "com.fasterxml.jackson.dataformat" % "jackson-dataformat-yaml" % "2.10.2",
  "com.fasterxml.jackson.module" %% "jackson-module-scala" % "2.10.2",
  // resource.managed
  "com.jsuereth" % "scala-arm_2.12" % "2.0"
  )
