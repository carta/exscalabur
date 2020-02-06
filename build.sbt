import sbt.Keys.libraryDependencies

name := "Exscalabur"
organization := "com.carta"
version := "0.0.1"

scalaVersion := "2.13.1"

libraryDependencies ++= Seq(
  "org.apache.poi" % "poi" % "4.1.0",
  "org.apache.poi" % "poi-ooxml" % "4.1.0",
  "org.scalactic" %% "scalactic" % "3.1.0",
  "org.scalatest" %% "scalatest" % "3.1.0" % "test",
)
