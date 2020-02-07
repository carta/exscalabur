import sbt.Keys.libraryDependencies


organization := "com.carta"
organizationName := "carta"

version := "0.0.3"
scalaVersion := "2.12.8"

name := "Exscalabur"
publishTo := { Some("Cloudsmith API" at sys.env.get("CLOUDSMITH_REPO").getOrElse("https://maven.cloudsmith.io/carta/maven-snapshots")) }

credentials += Credentials(
  "Cloudsmith API",
  "maven.cloudsmith.io",
  "token",
  sys.env.get("CLOUDSMITH_API_KEY").getOrElse("")
)

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
