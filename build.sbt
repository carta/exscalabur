import sbt.Keys.libraryDependencies


organization := "com.carta"
organizationName := "carta"

scalaVersion := "2.12.8"

val isSnapshot = sys.env.getOrElse("isSnapshot", "false").toBoolean
publishConfiguration := publishConfiguration.value.withOverwrite(isSnapshot)
name := "Exscalabur"
publishTo := {
  val baseUrl = "https://maven.cloudsmith.io/carta/";
  if (isSnapshot) {
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