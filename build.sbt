import sbt.Keys.libraryDependencies
import sbtrelease._
import sbtrelease.ReleaseStateTransformations.{setReleaseVersion=>_,_}


organization := "com.carta"
organizationName := "carta"

releaseUseGlobalVersion := false
scalaVersion := "2.12.8"

name := "Exscalabur"
publishTo := {
  val baseUrl = "https://maven.cloudsmith.io/carta/";
  if (isSnapshot.value)
    Some("Cloudsmith API Snapshots" at baseUrl + "maven-snapshots")
  else
    Some("Cloudsmith API Releases" at baseUrl + "maven-releases")
}

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

lazy val showVersion = taskKey[Unit]("Show version")
showVersion := {
  println(version.value)
}

git.useGitDescribe := true
git.formattedShaVersion := git.gitHeadCommit.value map { sha => s"v$sha" }


def setVersionOnly(selectVersion: Versions => String): ReleaseStep =  { st: State =>
  val vs = st.get(ReleaseKeys.versions).getOrElse(sys.error("No versions are set! Was this release part executed before inquireVersions?"))
  val selected = selectVersion(vs)

  st.log.info("Setting version to '%s'." format selected)
  val useGlobal =Project.extract(st).get(releaseUseGlobalVersion)
  val versionStr = (if (useGlobal) globalVersionString else versionString) format selected

  reapply(Seq(
    if (useGlobal) version in ThisBuild := selected
    else version := selected
  ), st)
}

lazy val setReleaseVersion: ReleaseStep = setVersionOnly(_._1)
