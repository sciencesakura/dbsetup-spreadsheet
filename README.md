# dbsetup-spreadsheet: Import Excel using DbSetup

[English](README.md) | [日本語](README.ja.md)

A [DbSetup](http://dbsetup.ninja-squad.com/) extension to import data from Microsoft Excel files.

![](https://github.com/sciencesakura/dbsetup-spreadsheet/workflows/build/badge.svg)

## Requirements

* Java 8+

## Installation

Gradle:

```groovy
testImplementation 'com.sciencesakura:dbsetup-spreadsheet:1.0.1'

// optional - Kotlin Extensions
testImplementation 'com.sciencesakura:dbsetup-spreadsheet-kt:1.0.1'

// optional - When import *.xlsx files
testRuntimeOnly 'org.apache.poi:poi-ooxml:5.1.0'
```

Maven:

```xml
<dependency>
  <groupId>com.sciencesakura</groupId>
  <artifactId>dbsetup-spreadsheet</artifactId>
  <version>1.0.1</version>
  <scope>test</scope>
</dependency>

<!-- optional - Kotlin Extensions -->
<dependency>
  <groupId>com.sciencesakura</groupId>
  <artifactId>dbsetup-spreadsheet-kt</artifactId>
  <version>1.0.1</version>
  <scope>test</scope>
</dependency>

<!-- optional - When import *.xlsx files -->
<dependency>
  <groupId>org.apache.poi</groupId>
  <artifactId>poi-ooxml</artifactId>
  <version>5.1.0</version>
  <scope>test</scope>
</dependency>
```

## Usage

When there are tables below:

```sql
create table country (
  id    integer       not null,
  code  char(3)       not null,
  name  varchar(256)  not null,
  primary key (id),
  unique (code)
);

create table customer (
  id      integer       not null,
  name    varchar(256)  not null,
  country integer       not null,
  primary key (id),
  foreign key (country) references country (id)
);
```

Create An Excel file with one worksheet per table, and name those worksheets the same name as the tables. If there is a dependency between the tables, put the worksheet of dependent table early.

`country` sheet:

|id|code|name|
|---:|---|---|
|1|GBR|United Kingdom|
|2|HKG|Hong Kong|
|3|JPN|Japan|

`customer` sheet:

|id|name|country|
|---:|---|---:|
|1|Eriol|1|
|2|Sakura|3|
|3|Xiaolang|2|

Put the prepared Excel file on the classpath, and write code like below.

Java:

```java
import static com.sciencesakura.dbsetup.spreadsheet.Import.excel;

Operation operation = excel("testdata.xlsx").build();
DbSetup dbSetup = new DbSetup(destination, operation);
dbSetup.launch();
```

Kotlin:

```kotlin
import com.sciencesakura.dbsetup.spreadsheet.excel

dbSetup(destination) {
    excel("testdata.xlsx")
}.launch()
```

See [API reference](https://sciencesakura.github.io/dbsetup-spreadsheet/) for details.

## Prefer CSV ?

→ [dbsetup-csv](https://github.com/sciencesakura/dbsetup-csv)

## License

MIT License

Copyright (c) 2019 sciencesakura
