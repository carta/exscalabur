release-snapshot:
	PROJECT_VERSION=$(sbt nextVersion | tail -2 | head -1) \
	COMMIT_HASH=$(git rev-parse HEAD) \
	sbt release release-version "$(PROJECT_VERSION)-$(COMMIT_HASH)" next-version $(PROJECT_VERSION)
