# dbsetup-spreadsheet: Import Excel using DbSetup

[English](README.md) | [日本語](README.ja.md)

Microsoft Excelファイルからデータ取り込みができる[DbSetup](http://dbsetup.ninja-squad.com/)拡張機能です.

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

次のテーブルがある場合:

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

テーブル毎に1枚のワークシートを含むExcelファイルを作成し, それらのワークシートにテーブルと同じ名前を付けます. テーブル間に依存関係がある場合は被依存テーブルのワークシートが先になるようにします.

`country` シート:

|id|code|name|
|---:|---|---|
|1|GBR|United Kingdom|
|2|HKG|Hong Kong|
|3|JPN|Japan|

`customer` シート:

|id|name|country|
|---:|---|---:|
|1|Eriol|1|
|2|Sakura|3|
|3|Xiaolang|2|

準備したExcelファイルをクラスパス上に置き, 次のようなコードを書きます:

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

詳細は[APIリファレンス](https://sciencesakura.github.io/dbsetup-spreadsheet/)を参照して下さい.

## Prefer CSV ?

→ [dbsetup-csv](https://github.com/sciencesakura/dbsetup-csv)

## License

MIT License

Copyright (c) 2019 sciencesakura
