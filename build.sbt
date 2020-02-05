/**
 * Copyright 2018 eShares, Inc. dba Carta, Inc.
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * https://www.apache.org/licenses/LICENSE-2.0
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
import sbt.Keys.libraryDependencies


organization := "com.carta"
organizationName := "carta"

scalaVersion := "2.12.8"

lazy val isSnapshotRelease = sys.env.getOrElse("isSnapshot", "false").toBoolean
publishConfiguration := publishConfiguration.value.withOverwrite(isSnapshotRelease)
name := "Exscalabur"
publishTo := {
  val baseUrl = "https://maven.cloudsmith.io/carta/";
  if (isSnapshotRelease) {
    Some("Cloudsmith API Snapshots" at baseUrl + "maven-snapshots")
  }
  else {
    Some("Cloudsmith API Releases" at baseUrl + "maven-releases")
  }
}

credentials += Credentials(
  "Cloudsmith API",
  "maven.cloudsmith.io",
  "token",
  sys.env.getOrElse("CLOUDSMITH_API_KEY", "")
)

libraryDependencies ++= Seq(
  "org.scalatest" %% "scalatest" % "3.1.0" % "test",
  "org.apache.poi" % "poi" % "4.1.0",
  "org.apache.poi" % "poi-ooxml" % "4.1.0",
  "com.fasterxml.jackson.core" % "jackson-databind" % "2.10.2",
  "com.fasterxml.jackson.dataformat" % "jackson-dataformat-yaml" % "2.10.2",
  "com.fasterxml.jackson.module" %% "jackson-module-scala" % "2.10.2",
  // resource.managed
  "com.jsuereth" % "scala-arm_2.12" % "2.0"
)

lazy val updateVersion = inputKey[Unit]("Updates version.sbt")
updateVersion := {
  import complete.DefaultParsers.spaceDelimited
  import java.io.FileOutputStream

  val versionIncrease = spaceDelimited("<arg>").parsed
    .headOption
    .filter(arg => arg == "MAJOR" || arg == "MINOR")
    .getOrElse("PATCH")

  val buildVersion = version.value.split(" ").last
  println(s"Release is $versionIncrease $buildVersion")

  val Array(major, minor, patch) = buildVersion.split("\\.").map(_.toInt)

  val nextVersion = (
    if (versionIncrease == "MAJOR") (major + 1, minor, patch)
    else if (versionIncrease == "MINOR") (major, minor + 1, patch)
    else if (versionIncrease == "PATCH") (major, minor, patch + 1)
    else (major, minor, patch)
    ).productIterator.mkString(".")

  val nextVersionLine = s"""version in ThisBuild := "$nextVersion""""
  println(s"Updating version.sbt to: $nextVersion")

  new FileOutputStream("./version.sbt").write(nextVersionLine.getBytes)
  println(s"Success!")
}