name := "demo"
version := "0.1"
scalaVersion := "2.12.8"
organization := "com.carta"
organizationName := "carta"

resolvers += "carta-maven-releases" at "https://dl.cloudsmith.io/xflWv9geX7qnfq1J/carta/maven-releases/maven/"

libraryDependencies ++= Seq(
  "com.carta" %% "exscalabur" % "0.0.8"
)