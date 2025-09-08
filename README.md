# dbsetup-spreadsheet: Import Excel into database with DbSetup

![](https://github.com/sciencesakura/dbsetup-spreadsheet/actions/workflows/build.yaml/badge.svg) [![Maven Central](https://maven-badges.sml.io/sonatype-central/com.sciencesakura/dbsetup-spreadsheet/badge.svg)](https://maven-badges.sml.io/sonatype-central/com.sciencesakura/dbsetup-spreadsheet)

A [DbSetup](http://dbsetup.ninja-squad.com/) extension for importing Excel files into the database.

## Requirements

* Java 11+

## Installation

The dbsetup-spreadsheet library is available on Maven Central. You can install it using your build system of choice.

### Gradle

```groovy
testImplementation 'com.sciencesakura:dbsetup-spreadsheet:2.0.2'
```

If you are using Kotlin, you can use the Kotlin module for a more concise DSL:

```groovy
testImplementation 'com.sciencesakura:dbsetup-spreadsheet-kt:2.0.2'
```

### Maven

```xml
<dependency>
  <groupId>com.sciencesakura</groupId>
  <artifactId>dbsetup-spreadsheet</artifactId>
  <version>2.0.2</version>
  <scope>test</scope>
</dependency>
```

If you are using Kotlin, you can use the Kotlin module for a more concise DSL:

```xml
<dependency>
  <groupId>com.sciencesakura</groupId>
  <artifactId>dbsetup-spreadsheet-kt</artifactId>
  <version>2.0.2</version>
  <scope>test</scope>
</dependency>
```

## Usage

### Import Excel file

When the following two tables exist in your database:

```sql
create table countries (
  id    integer      not null,
  code  char(3)      not null,
  name  varchar(256) not null,
  primary key (id),
  unique (code)
);

create table customers (
  id      integer      not null,
  name    varchar(256) not null,
  country integer      not null,
  primary key (id),
  foreign key (country) references countries (id)
);
```

You can import an Excel file into the above two tables as follows:

```java
import static com.sciencesakura.dbsetup.spreadsheet.Import.excel;
import com.ninja_squad.dbsetup.DbSetup;

// The operation to import an Excel file into a tables
var operation = excel("test-data.xlsx").build();

// Create a `DbSetup` instance with the operation and execute it
var dbSetup = new DbSetup(destination, operation);
dbSetup.launch();
```

The Excel file `test-data.xlsx` is composed of two worksheets:

**`countries` worksheet**

| id | code | name           |
|----|------|----------------|
| 1  | GBR  | United Kingdom |
| 2  | HKG  | Hong Kong      |
| 3  | JPN  | Japan          |

**`customers` worksheet**

| id | name     | country |
|----|----------|---------|
| 1  | Eriol    | 1       |
| 2  | Sakura   | 3       |
| 3  | Xiaolang | 2       |

**Note:** There are dependencies between the two tables, so the `countries` worksheet must be before the `customers` worksheet in the Excel file.

### Exclude worksheets from importing

```java
import static com.sciencesakura.dbsetup.spreadsheet.Import.excel;

var operation = excel("test-data.xlsx")
    // Do not import `README` worksheet and any worksheet whose name starts with `temp-`
    .exclude("README", "^temp-.+")
    .build();
```

### Customize mapping of worksheet names to table names

```java
import static com.sciencesakura.dbsetup.spreadsheet.Import.excel;

var operation = excel("test-data.xlsx")
    // Map the `data-xxx` worksheet to the `xxx` table
    .resolver(sht -> sht.replaceFirst("^data-", ""))
    .build();
```

### Clear table before import

```java
import static com.ninja_squad.dbsetup.Operations.*;
import static com.sciencesakura.dbsetup.spreadsheet.Import.excel;
import com.ninja_squad.dbsetup.DbSetup;

// The operations to clear the tables and then import the Excel file
var operation = sequenceOf(
    deleteAllFrom("customers", "countries"),
    excel("test-data.xlsx").build()
);

var dbSetup = new DbSetup(destination, operation);
dbSetup.launch();
```

### Use generated values and fixed values

```java
import static com.sciencesakura.dbsetup.spreadsheet.Import.excel;
import com.ninja_squad.dbsetup.generator.ValueGenerators;

var operation = excel("test-data.xlsx")
    // Generate a random UUID for the `items` table's `id` column
    .withGeneratedValue("items", "id", () -> UUID.randomUUID().toString())
    // Generate a sequential string for the `items` table's `name` column, starting with "item-001"
    .withGeneratedValue("items", "name", ValueGenerators.stringSequence("item-").withLeftPadding(3))
    // Set a fixed value for the `items` table's `created_at` column
    .withDefaultValue("created_at", "2023-01-01 10:20:30")
    .build();
```

### Use Kotlin DSL

```kotlin
import com.ninja_squad.dbsetup_kotlin.dbSetup
import com.sciencesakura.dbsetup.spreadsheet.excel

dbSetup(destination) {
  excel("test-data.xlsx") {
    exclude("README")
    withGeneratedValue("items", "id") { UUID.randomUUID().toString() }
  }
}.launch()
```

See [API reference](https://sciencesakura.github.io/dbsetup-spreadsheet/) for more details.

## Prefer CSV?

→ [dbsetup-csv](https://github.com/sciencesakura/dbsetup-csv)

## License

This library is licensed under the MIT License.

Copyright (c) 2019 sciencesakura
