# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]
### Changed
- ci: Set paths to reduce unnecessary CI runs (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/59).
- build: Upgrade the Java used for building from 17 to 21 (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/60).
- build: Add Dependabot (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/61).
- ci: Deploy docs automatically (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/63).
- Bump org.junit.jupiter:junit-jupiter from 6.0.3 to 6.1.2 (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/62).
- Bump org.apache.logging.log4j:log4j-core from 2.25.3 to 2.25.4 (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/58).
- Bump kotlin.version from 2.3.10 to 2.4.10 (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/64).
- build: Remove the redundant option kotlin.compiler.jvmTarget (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/69).
- Bump org.jspecify:jspecify from 1.0.0 to 1.0.1 (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/65).
- Bump org.assertj:assertj-db from 3.0.1 to 3.0.2 (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/66).
- build: Follow the maven-bundle-plugin's default behavior for Import-Package (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/70).
- Bump org.apache.maven.plugins:maven-enforcer-plugin from 3.6.2 to 3.6.3 (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/67).
- Bump org.apache.felix:maven-bundle-plugin from 6.0.0 to 6.1.0 (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/68).
- feat: Use package-level nullability annotation (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/71).

## [2.0.3] - 2026-02-22
### Changed
- Upgrade tool Java version to 17 (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/54).
- Upgrade dependencies (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/55).
- Switch to JSpecify for nullability expression (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/56).

## [2.0.2] - 2025-08-15
### Changed
- Upgrade toolchains and improve dev environment (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/50).
- Various corrections (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/51).

## [2.0.1] - 2024-03-23
### Added
- Allow ignorable rows between the header row and the data rows (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/47).
### Changed
- Upgrade POI 5.2.3->5.2.5 and others (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/45).
- Upgrade actions: checkout and setup-java (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/46).

## [2.0.0] - 2023-07-30
### Changed
- Build with maven (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/38).
- Make it possible to publish API docs with CI (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/39).
- Publish packages to registry in CI (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/40).
- Improve test (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/41).
- Refactor (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/42).
- Change and improve tableMapper method (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/43).

## [1.0.3] - 2022-06-01
### Changed
- Improve 'Installation' section of README.md (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/33).
- Upgrade POI 5.1.0->5.2.2 and others (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/34).
- Improve sheet exclusion (https://github.com/sciencesakura/dbsetup-spreadsheet/pull/35).
