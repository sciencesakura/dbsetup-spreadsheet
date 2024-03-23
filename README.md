# dbsetup-spreadsheet: Import Excel using DbSetup

[English](README.md) | [日本語](README.ja.md)

A [DbSetup](http://dbsetup.ninja-squad.com/) extension to import data from Microsoft Excel files.

![](https://github.com/sciencesakura/dbsetup-spreadsheet/actions/workflows/check.yaml/badge.svg) [![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.sciencesakura/dbsetup-spreadsheet/badge.svg)](https://maven-badges.herokuapp.com/maven-central/com.sciencesakura/dbsetup-spreadsheet)

## Requirements

* Java 11+

## Installation

### Gradle

#### Java

```groovy
testImplementation 'com.sciencesakura:dbsetup-spreadsheet:2.0.1'
```

#### Kotlin

```groovy
testImplementation 'com.sciencesakura:dbsetup-spreadsheet-kt:2.0.1'
```

### Maven

#### Java

```xml
<dependency>
  <groupId>com.sciencesakura</groupId>
  <artifactId>dbsetup-spreadsheet</artifactId>
  <version>2.0.1</version>
  <scope>test</scope>
</dependency>
```

#### Kotlin

```xml
<dependency>
  <groupId>com.sciencesakura</groupId>
  <artifactId>dbsetup-spreadsheet-kt</artifactId>
  <version>2.0.1</version>
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

Create An Excel file with one worksheet per table, and name those worksheets the same as the tables. If there is a dependency between the tables, put the worksheet of the dependent table early.

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

Put the prepared Excel file on the classpath, and write code like the below:

Java:

```java
import static com.sciencesakura.dbsetup.spreadsheet.Import.excel;

var operation = excel("testdata.xlsx").build();
var dbSetup = new DbSetup(destination, operation);
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
