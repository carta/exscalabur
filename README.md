[![CircleCI](https://circleci.com/gh/carta/exscalabur/tree/master.svg?style=svg&circle-token=a8e8f68d2e70a177a3298140e5ec935710f651c7)](https://circleci.com/gh/carta/exscalabur/tree/master)
[![Scala version](https://img.shields.io/badge/scala-2.12.8-brightgreen.svg)](https://www.scala-lang.org/download/2.12.8.html)
[![Latest Version @ Cloudsmith](https://api-prd.cloudsmith.io/badges/version/carta/maven-releases/maven/exscalabur_2.12/latest/xg=com.carta/?render=true&badge_token=gAAAAABePJwL9tZMOa6DrXC6N_iGkYROA2I1jTSwarIRvAuhy7O34Tt742-doost6rUHEs5WR2PqRoxjGCihc1v0mCeHIeVY_hSi6-wyPttrjUAFaGPmXMU%3D)](https://cloudsmith.io/~carta/repos/maven-releases/packages/detail/maven/exscalabur_2.12/latest/xg=com.carta/)

# Exscalabur
A scala package for generating excel exports from an excel template and JSON input.

## Usage
Add it as a dependency:
`libraryDependencies += "com.carta" % "exscalabur" % "2.12.8"`

## Development
1. `brew install sbt`
2. `sbt` (https://www.scala-sbt.org/1.x/docs/)

### Common sbt Commands

| Command | Description | 
|---|---|
| `~compile` | `~` enables hot reloading |
| `~run` | `~` enables hot reloading |
| `test` | Runs all tests |
| `testQuick` | Runs tests only affected by your latest code changes |
| `clean` | Removes generated files from the target directory |
| `update` | Updates external dependencies |
| `package` | Creates JAR |
| `publishLocal` | Publishes to local IVY repository, by default in `~/.ivy2/local`
