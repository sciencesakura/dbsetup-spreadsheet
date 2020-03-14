# dbsetup-spreadsheet: Import Excel using DbSetup

[English](README.md) | [日本語](README.ja.md)

A [DbSetup](http://dbsetup.ninja-squad.com/) extension to import data from Microsoft Excel files.

![](https://github.com/sciencesakura/dbsetup-spreadsheet/workflows/build/badge.svg)

## Requirement

* Java 8 or later

## Installation

Gradle:

```groovy
testImplementation 'com.sciencesakura:dbsetup-spreadsheet:0.0.2'
testRuntimeOnly 'org.apache.poi:poi-ooxml:4.1.1' // if import *.xlsx
```

Maven:

```xml
<dependency>
  <groupId>com.sciencesakura</groupId>
  <artifactId>dbsetup-spreadsheet</artifactId>
  <version>0.0.2</version>
  <scope>test</scope>
</dependency>
<dependency><!-- if import *.xlsx -->
  <groupId>org.apache.poi</groupId>
  <artifactId>poi-ooxml</artifactId>
  <version>4.1.1</version>
  <scope>test</scope>
</dependency>
```

## Usage

```java
import static com.sciencesakura.dbsetup.spreadsheet.Import.excel;

// `testdata.xlsx` must be in classpath.
Operation operation = excel("testdata.xlsx").build();
DbSetup dbSetup = new DbSetup(destination, operation);
dbSetup.launch();
```

See [API reference](https://sciencesakura.github.io/dbsetup-spreadsheet/) for details.

## Recommendation

We recommend using this extension only when the destination table has too many columns to keep your code using the [Insert.Builder](http://dbsetup.ninja-squad.com/apidoc/2.1.0/com/ninja_squad/dbsetup/operation/Insert.Builder.html) class readable.

## Prefer CSV ?

→ [dbsetup-csv](https://github.com/sciencesakura/dbsetup-csv)

## License

MIT License

Copyright (c) 2019 sciencesakura
