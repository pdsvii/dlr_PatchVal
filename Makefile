# Makefile for Patch Validator
.PHONY: changelog-from-git changelog-release

changelog-from-git:
	python scripts/generate_changelog.py --from-git

# Usage: make changelog-release version=0.1.0
changelog-release:
	python scripts/generate_changelog.py --release $(version)
