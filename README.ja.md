# dbsetup-spreadsheet: Import Excel using DbSetup

[English](README.md) | [日本語](README.ja.md)

Microsoft Excelファイルからデータ取り込みができる[DbSetup](http://dbsetup.ninja-squad.com/)拡張機能です.

## Requirement

* Java 8以降

## Installation

Gradle:

```groovy
testImplementation 'com.sciencesakura:dbsetup-spreadsheet:0.0.2'
testRuntimeOnly 'org.apache.poi:poi-ooxml:4.1.1' // *.xlsxをインポートする場合
```

Maven:

```xml
<dependency>
  <groupId>com.sciencesakura</groupId>
  <artifactId>dbsetup-spreadsheet</artifactId>
  <version>0.0.2</version>
  <scope>test</scope>
</dependency>
<dependency><!-- *.xlsxをインポートする場合 -->
  <groupId>org.apache.poi</groupId>
  <artifactId>poi-ooxml</artifactId>
  <version>4.1.1</version>
  <scope>test</scope>
</dependency>
```

## Usage

```java
import static com.sciencesakura.dbsetup.spreadsheet.Import.excel;

// `testdata.xlsx` はクラスパス上にある必要があります
Operation operation = excel("testdata.xlsx").build();
DbSetup dbSetup = new DbSetup(destination, operation);
dbSetup.launch();
```

詳細は[APIリファレンス](https://sciencesakura.github.io/dbsetup-spreadsheet/)を参照して下さい.

## Recommendation

この拡張機能を利用するのは, 取り込み先テーブルの列数が多すぎて[Insert.Builder](http://dbsetup.ninja-squad.com/apidoc/2.1.0/com/ninja_squad/dbsetup/operation/Insert.Builder.html)を使用したコードの可読性が悪くなってしまう場合にのみにすることをお薦めします.

## Prefer CSV ?

→ [dbsetup-csv](https://github.com/sciencesakura/dbsetup-csv)

## License

MIT License

Copyright (c) 2019 sciencesakura
