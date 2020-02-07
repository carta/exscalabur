import sbt.Keys.libraryDependencies
import sbtrelease._
import sbtrelease.ReleaseStateTransformations.{setReleaseVersion=>_,_}
import ReleaseTransformations._


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

lazy val nextVersion = taskKey[Unit]("Show next version")
nextVersion := {
  val predicted = Version(version.value).map(_.bump(releaseVersionBump.value).asSnapshot.string).getOrElse(versionFormatError(version.value));
  println(predicted)
}

git.useGitDescribe := true
git.formattedShaVersion := git.gitHeadCommit.value map { sha => s"v$sha" }


releaseProcess := {
  if (isSnapshot.value)
    // Release without git steps
    Seq[ReleaseStep](
      checkSnapshotDependencies,
      inquireVersions,
      runTest,
      setReleaseVersion,
      publishArtifacts,
      setNextVersion
    )
  else
    Seq[ReleaseStep](
      checkSnapshotDependencies,              // : ReleaseStep
      inquireVersions,                        // : ReleaseStep
      runClean,                               // : ReleaseStep
      runTest,                                // : ReleaseStep
      setReleaseVersion,                      // : ReleaseStep
      commitReleaseVersion,                   // : ReleaseStep, performs the initial git checks
      tagRelease,                             // : ReleaseStep
      publishArtifacts,                       // : ReleaseStep, checks whether `publishTo` is properly set up
      setNextVersion,                         // : ReleaseStep
      commitNextVersion,                      // : ReleaseStep
      pushChanges                             // : ReleaseStep, also checks that an upstream branch is properly configured
    )
}