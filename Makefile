PROJECT_VERSION := `sbt nextVersion | tail -2 | head -1`
COMMIT_HASH := `git rev-parse HEAD`

release-snapshot:
	@echo "Attempting release of ${PROJECT_VERSION}-${COMMIT_HASH}"
	@sbt release release-version "${PROJECT_VERSION}-${COMMIT_HASH}" next-version $(PROJECT_VERSION)
