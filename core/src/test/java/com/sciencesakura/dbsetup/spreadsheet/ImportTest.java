// SPDX-License-Identifier: MIT

package com.sciencesakura.dbsetup.spreadsheet;

import static com.ninja_squad.dbsetup.Operations.sequenceOf;
import static com.ninja_squad.dbsetup.Operations.sql;
import static com.ninja_squad.dbsetup.Operations.truncate;
import static com.sciencesakura.dbsetup.spreadsheet.Import.excel;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.assertj.db.api.Assertions.assertThat;

import com.ninja_squad.dbsetup.DbSetup;
import com.ninja_squad.dbsetup.DbSetupRuntimeException;
import com.ninja_squad.dbsetup.destination.Destination;
import com.ninja_squad.dbsetup.destination.DriverManagerDestination;
import com.ninja_squad.dbsetup.generator.ValueGenerators;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Map;
import java.util.UUID;
import java.util.function.Function;
import java.util.regex.Pattern;
import org.assertj.db.type.AssertDbConnection;
import org.assertj.db.type.AssertDbConnectionFactory;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

class ImportTest {

  static final String url = "jdbc:h2:mem:test;DB_CLOSE_DELAY=-1";

  static final String username = "sa";

  AssertDbConnection connection;

  Destination destination;

  @BeforeEach
  void setUp() {
    connection = AssertDbConnectionFactory.of(url, username, null).create();
    destination = new DriverManagerDestination(url, username, null);
  }

  @Nested
  class DataTypes {

    @BeforeEach
    void setUp() {
      var ddl = sql("create table if not exists data_types ("
          + "id uuid not null,"
          + "num1 smallint,"
          + "num2 integer,"
          + "num3 bigint,"
          + "num4 real,"
          + "num5 decimal(7,3),"
          + "text1 char(5),"
          + "text2 varchar(100),"
          + "date1 timestamp,"
          + "date2 date,"
          + "date3 time,"
          + "bool1 boolean,"
          + "primary key (id)"
          + ")");
      new DbSetup(destination, sequenceOf(ddl, truncate("data_types"))).launch();
    }

    @Test
    void import_with_default_settings() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("DataTypes/data_types.xlsx").build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(3)
          .changeOfCreation()
          .rowAtEndPoint()
          .value("id").isEqualTo(new UUID(0, 1))
          .value("num1").isEqualTo(1000)
          .value("num2").isEqualTo(20000)
          .value("num3").isEqualTo(3000000000L)
          .value("num4").isEqualTo(400.75)
          .value("num5").isEqualTo(new BigDecimal("5000.333"))
          .value("text1").isEqualTo("aaa  ")
          .value("text2").isEqualTo("bbb")
          .value("date1").isEqualTo(LocalDateTime.parse("2001-02-03T10:20:30.456"))
          .value("date2").isEqualTo(LocalDate.parse("2001-02-03"))
          .value("date3").isEqualTo(LocalTime.parse("10:20:30"))
          .value("bool1").isTrue()
          .changeOfCreation()
          .rowAtEndPoint()
          .value("id").isEqualTo(new UUID(0, 2))
          .value("num1").isNull()
          .value("num2").isNull()
          .value("num3").isNull()
          .value("num4").isNull()
          .value("num5").isNull()
          .value("text1").isNull()
          .value("text2").isNull()
          .value("date1").isNull()
          .value("date2").isNull()
          .value("date3").isNull()
          .value("bool1").isFalse()
          .changeOfCreation()
          .rowAtEndPoint()
          .value("id").isEqualTo(new UUID(0, 3))
          .value("num1").isEqualTo(1001)
          .value("num2").isEqualTo(20001)
          .value("num3").isEqualTo(3000000001L)
          .value("num4").isEqualTo(401.75)
          .value("num5").isEqualTo(new BigDecimal("5001.333"))
          .value("text1").isNull()
          .value("text2").isEqualTo("aaabbb")
          .value("date1").isEqualTo(LocalDateTime.parse("2001-02-04T10:20:30.456"))
          .value("date2").isEqualTo(LocalDate.parse("2001-02-04"))
          .value("date3").isNull()
          .value("bool1").isTrue();
    }
  }

  @Nested
  @SuppressWarnings("ConstantConditions")
  class ExcelFile {

    @Test
    void throw_npe_if_location_is_null() {
      assertThatThrownBy(() -> excel(null))
          .isInstanceOf(NullPointerException.class)
          .hasMessage("location must not be null");
    }

    @Test
    void throw_iae_if_location_has_bean_not_found() {
      assertThatThrownBy(() -> excel("ExcelFile/not_found.xlsx"))
          .isInstanceOf(IllegalArgumentException.class)
          .hasMessage("ExcelFile/not_found.xlsx not found");
    }

    @Test
    void throw_dsre_if_header_row_contains_blank() {
      assertThatThrownBy(() -> excel("ExcelFile/invalid_sheets.xlsx").include("blank_in_header").build())
          .isInstanceOf(DbSetupRuntimeException.class)
          .hasMessage("header cell must not be blank: blank_in_header!B1");
    }

    @Test
    void throw_dsre_if_header_row_contains_non_string_value() {
      assertThatThrownBy(() -> excel("ExcelFile/invalid_sheets.xlsx").include("non_string_in_header").build())
          .isInstanceOf(DbSetupRuntimeException.class)
          .hasMessage("header cell must be string type: non_string_in_header!B1");
    }

    @Test
    void throw_dsre_if_header_row_contains_error() {
      assertThatThrownBy(() -> excel("ExcelFile/invalid_sheets.xlsx").include("error_in_header").build())
          .isInstanceOf(DbSetupRuntimeException.class)
          .hasMessage("error value contained: error_in_header!B1");
    }
  }

  @Nested
  @SuppressWarnings("ConstantConditions")
  class TableNames {

    @BeforeEach
    void setUp() {
      var table_11 = "create table if not exists table_11 ("
          + "id integer primary key,"
          + "name varchar(100)"
          + ")";
      var table_12 = "create table if not exists table_12 ("
          + "id integer primary key,"
          + "name varchar(100)"
          + ")";
      var table_21 = "create table if not exists table_21 ("
          + "id integer primary key,"
          + "name varchar(100)"
          + ")";
      var table_22 = "create table if not exists table_22 ("
          + "id integer primary key,"
          + "name varchar(100)"
          + ")";
      new DbSetup(destination, sequenceOf(sql(table_11, table_12, table_21, table_22),
          truncate("table_11", "table_12", "table_21", "table_22"))).launch();
    }

    @Test
    void import_all_sheets() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("TableNames/table_names.xlsx").build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(4)
          .changeOfCreationOnTable("table_11")
          .rowAtEndPoint()
          .value("id").isEqualTo(1)
          .value("name").isEqualTo("Alice")
          .changeOfCreationOnTable("table_12")
          .rowAtEndPoint()
          .value("id").isEqualTo(2)
          .value("name").isEqualTo("Bob")
          .changeOfCreationOnTable("table_21")
          .rowAtEndPoint()
          .value("id").isEqualTo(3)
          .value("name").isEqualTo("Charlie")
          .changeOfCreationOnTable("table_22")
          .rowAtEndPoint()
          .value("id").isEqualTo(4)
          .value("name").isEqualTo("Dave");
    }

    @Test
    void import_only_sheets_whose_name_matches_string_pattern() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("TableNames/table_names.xlsx")
          .include(".+2$").build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(2)
          .changeOfCreationOnTable("table_12")
          .rowAtEndPoint()
          .value("id").isEqualTo(2)
          .value("name").isEqualTo("Bob")
          .changeOfCreationOnTable("table_22")
          .rowAtEndPoint()
          .value("id").isEqualTo(4)
          .value("name").isEqualTo("Dave");
    }

    @Test
    void import_only_sheets_whose_name_matches_string_patterns() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("TableNames/table_names.xlsx")
          .include(".+2$", ".+11$").build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(3)
          .changeOfCreationOnTable("table_11")
          .rowAtEndPoint()
          .value("id").isEqualTo(1)
          .value("name").isEqualTo("Alice")
          .changeOfCreationOnTable("table_12")
          .rowAtEndPoint()
          .value("id").isEqualTo(2)
          .value("name").isEqualTo("Bob")
          .changeOfCreationOnTable("table_22")
          .rowAtEndPoint()
          .value("id").isEqualTo(4)
          .value("name").isEqualTo("Dave");
    }

    @Test
    void import_only_sheets_whose_name_matches_pattern() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("TableNames/table_names.xlsx")
          .include(Pattern.compile(".+2$")).build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(2)
          .changeOfCreationOnTable("table_12")
          .rowAtEndPoint()
          .value("id").isEqualTo(2)
          .value("name").isEqualTo("Bob")
          .changeOfCreationOnTable("table_22")
          .rowAtEndPoint()
          .value("id").isEqualTo(4)
          .value("name").isEqualTo("Dave");
    }

    @Test
    void import_only_sheets_whose_name_matches_patterns() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("TableNames/table_names.xlsx")
          .include(Pattern.compile(".+2$"), Pattern.compile(".+11$")).build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(3)
          .changeOfCreationOnTable("table_11")
          .rowAtEndPoint()
          .value("id").isEqualTo(1)
          .value("name").isEqualTo("Alice")
          .changeOfCreationOnTable("table_12")
          .rowAtEndPoint()
          .value("id").isEqualTo(2)
          .value("name").isEqualTo("Bob")
          .changeOfCreationOnTable("table_22")
          .rowAtEndPoint()
          .value("id").isEqualTo(4)
          .value("name").isEqualTo("Dave");
    }

    @Test
    void skip_sheets_with_name_that_matches_string_pattern() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("TableNames/table_names.xlsx")
          .exclude(".+2$").build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(2)
          .changeOfCreationOnTable("table_11")
          .rowAtEndPoint()
          .value("id").isEqualTo(1)
          .value("name").isEqualTo("Alice")
          .changeOfCreationOnTable("table_21")
          .rowAtEndPoint()
          .value("id").isEqualTo(3)
          .value("name").isEqualTo("Charlie");
    }

    @Test
    void skip_sheets_with_name_that_matches_string_patterns() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("TableNames/table_names.xlsx")
          .exclude(".+2$", ".+11$").build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(1)
          .changeOfCreationOnTable("table_21")
          .rowAtEndPoint()
          .value("id").isEqualTo(3)
          .value("name").isEqualTo("Charlie");
    }

    @Test
    void skip_sheets_with_name_that_matches_pattern() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("TableNames/table_names.xlsx")
          .exclude(Pattern.compile(".+2$")).build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(2)
          .changeOfCreationOnTable("table_11")
          .rowAtEndPoint()
          .value("id").isEqualTo(1)
          .value("name").isEqualTo("Alice")
          .changeOfCreationOnTable("table_21")
          .rowAtEndPoint()
          .value("id").isEqualTo(3)
          .value("name").isEqualTo("Charlie");
    }

    @Test
    void skip_sheets_with_name_that_matches_patterns() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("TableNames/table_names.xlsx")
          .exclude(Pattern.compile(".+2$"), Pattern.compile(".+11$")).build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(1)
          .changeOfCreationOnTable("table_21")
          .rowAtEndPoint()
          .value("id").isEqualTo(3)
          .value("name").isEqualTo("Charlie");
    }

    @Test
    void throws_npe_if_string_pattern_to_include_is_null() {
      String[] patterns = null;
      assertThatThrownBy(() -> excel("TableNames/table_names.xlsx").include(patterns))
          .isInstanceOf(NullPointerException.class)
          .hasMessage("patterns must not be null");
    }

    @Test
    void throws_npe_if_string_pattern_to_include_contains_null() {
      String[] patterns = new String[]{".+2$", null};
      assertThatThrownBy(() -> excel("TableNames/table_names.xlsx").include(patterns))
          .isInstanceOf(NullPointerException.class)
          .hasMessage("patterns must not contain null");
    }

    @Test
    void throws_npe_if_pattern_to_include_is_null() {
      Pattern[] patterns = null;
      assertThatThrownBy(() -> excel("TableNames/table_names.xlsx").include(patterns))
          .isInstanceOf(NullPointerException.class)
          .hasMessage("patterns must not be null");
    }

    @Test
    void throws_npe_if_pattern_to_include_contains_null() {
      Pattern[] patterns = new Pattern[]{Pattern.compile(".+2$"), null};
      assertThatThrownBy(() -> excel("TableNames/table_names.xlsx").include(patterns))
          .isInstanceOf(NullPointerException.class)
          .hasMessage("patterns must not contain null");
    }
  }

  @Nested
  @SuppressWarnings("ConstantConditions")
  class TableMapping {

    @BeforeEach
    void setUp() {
      var table_11 = "create table if not exists table_11 ("
          + "id integer primary key,"
          + "name varchar(100)"
          + ")";
      var table_12 = "create table if not exists table_12 ("
          + "id integer primary key,"
          + "name varchar(100)"
          + ")";
      var table_13 = "create table if not exists table_13 ("
          + "id integer primary key,"
          + "name varchar(100)"
          + ")";
      new DbSetup(destination, sequenceOf(sql(table_11, table_12, table_13),
          truncate("table_11", "table_12", "table_13"))).launch();
    }

    @Test
    void sheet_name_maps_to_table_name() {
      var changes = connection.changes().build().setStartPointNow();
      var resolver = Map.of("a", "table_13", "b", "table_12", "c", "table_11");
      var operation = excel("TableMapping/table_mapping.xlsx")
          .resolver(resolver).build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(3)
          .changeOfCreationOnTable("table_11")
          .rowAtEndPoint()
          .value("id").isEqualTo(3)
          .value("name").isEqualTo("Charlie")
          .changeOfCreationOnTable("table_12")
          .rowAtEndPoint()
          .value("id").isEqualTo(2)
          .value("name").isEqualTo("Bob")
          .changeOfCreationOnTable("table_13")
          .rowAtEndPoint()
          .value("id").isEqualTo(1)
          .value("name").isEqualTo("Alice");
    }

    @Test
    void throws_dsre_if_resolver_could_not_resolve_table() {
      var resolver = Map.of("a", "table_13", "b", "table_12");
      assertThatThrownBy(() -> excel("TableMapping/table_mapping.xlsx").resolver(resolver).build())
          .isInstanceOf(DbSetupRuntimeException.class)
          .hasMessage("could not resolve table name: c");
    }

    @Test
    void throws_npe_if_resolver_is_null_1() {
      assertThatThrownBy(() -> excel("TableNames/table_names.xlsx").resolver((Map<String, String>) null))
          .isInstanceOf(NullPointerException.class)
          .hasMessage("resolver must not be null");
    }

    @Test
    void throws_npe_if_resolver_is_null_2() {
      assertThatThrownBy(() -> excel("TableNames/table_names.xlsx").resolver((Function<String, String>) null))
          .isInstanceOf(NullPointerException.class)
          .hasMessage("resolver must not be null");
    }
  }

  @Nested
  class Margin {

    @BeforeEach
    void setUp() {
      var table_11 = "create table if not exists table_11 ("
          + "id integer primary key,"
          + "name varchar(100)"
          + ")";
      var table_12 = "create table if not exists table_12 ("
          + "id integer primary key,"
          + "name varchar(100)"
          + ")";
      new DbSetup(destination, sequenceOf(sql(table_11, table_12),
          truncate("table_11", "table_12"))).launch();
    }

    @Test
    void default_margin_is_zero() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("Margin/no_margin.xlsx").build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(2)
          .changeOfCreationOnTable("table_11")
          .rowAtEndPoint()
          .value("id").isEqualTo(1)
          .value("name").isEqualTo("Alice")
          .changeOfCreationOnTable("table_12")
          .rowAtEndPoint()
          .value("id").isEqualTo(2)
          .value("name").isEqualTo("Bob");
    }

    @Test
    void left_is_zero() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("Margin/no_margin.xlsx").left(0).build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(2)
          .changeOfCreationOnTable("table_11")
          .rowAtEndPoint()
          .value("id").isEqualTo(1)
          .value("name").isEqualTo("Alice")
          .changeOfCreationOnTable("table_12")
          .rowAtEndPoint()
          .value("id").isEqualTo(2)
          .value("name").isEqualTo("Bob");
    }

    @Test
    void left_is_five() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("Margin/left_margin.xlsx").left(5).build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(2)
          .changeOfCreationOnTable("table_11")
          .rowAtEndPoint()
          .value("id").isEqualTo(1)
          .value("name").isEqualTo("Alice")
          .changeOfCreationOnTable("table_12")
          .rowAtEndPoint()
          .value("id").isEqualTo(2)
          .value("name").isEqualTo("Bob");
    }

    @Test
    void top_is_zero() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("Margin/no_margin.xlsx").top(0).build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(2)
          .changeOfCreationOnTable("table_11")
          .rowAtEndPoint()
          .value("id").isEqualTo(1)
          .value("name").isEqualTo("Alice")
          .changeOfCreationOnTable("table_12")
          .rowAtEndPoint()
          .value("id").isEqualTo(2)
          .value("name").isEqualTo("Bob");
    }

    @Test
    void top_is_five() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("Margin/top_margin.xlsx").top(5).build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(2)
          .changeOfCreationOnTable("table_11")
          .rowAtEndPoint()
          .value("id").isEqualTo(1)
          .value("name").isEqualTo("Alice")
          .changeOfCreationOnTable("table_12")
          .rowAtEndPoint()
          .value("id").isEqualTo(2)
          .value("name").isEqualTo("Bob");
    }

    @Test
    void margin_is_zero() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("Margin/no_margin.xlsx").margin(0, 0).build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(2)
          .changeOfCreationOnTable("table_11")
          .rowAtEndPoint()
          .value("id").isEqualTo(1)
          .value("name").isEqualTo("Alice")
          .changeOfCreationOnTable("table_12")
          .rowAtEndPoint()
          .value("id").isEqualTo(2)
          .value("name").isEqualTo("Bob");
    }

    @Test
    void left_is_two_and_top_is_three() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("Margin/left_top_margin.xlsx").margin(2, 3).build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(2)
          .changeOfCreationOnTable("table_11")
          .rowAtEndPoint()
          .value("id").isEqualTo(1)
          .value("name").isEqualTo("Alice")
          .changeOfCreationOnTable("table_12")
          .rowAtEndPoint()
          .value("id").isEqualTo(2)
          .value("name").isEqualTo("Bob");
    }

    @Test
    void margin_between_header_and_data() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("Margin/margin_between_header_and_data.xlsx").skipAfterHeader(1).build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(2)
          .changeOfCreationOnTable("table_11")
          .rowAtEndPoint()
          .value("id").isEqualTo(1)
          .value("name").isEqualTo("Alice")
          .changeOfCreationOnTable("table_12")
          .rowAtEndPoint()
          .value("id").isEqualTo(2)
          .value("name").isEqualTo("Bob");
    }

    @Test
    void throw_dsre_if_header_row_is_not_found_1() {
      assertThatThrownBy(() -> excel("Margin/no_margin.xlsx").top(8).build())
          .isInstanceOf(DbSetupRuntimeException.class)
          .hasMessage("header row not found: table_11[8]");
    }

    @Test
    void throw_dsre_if_header_row_is_not_found_2() {
      assertThatThrownBy(() -> excel("Margin/no_margin.xlsx").left(8).build())
          .isInstanceOf(DbSetupRuntimeException.class)
          .hasMessage("header row not found: table_11[0]");
    }

    @Test
    void throws_iae_if_left_is_negative() {
      assertThatThrownBy(() -> excel("Margin/no_margin.xlsx").left(-1))
          .isInstanceOf(IllegalArgumentException.class)
          .hasMessage("left must be greater than or equal to 0");
    }

    @Test
    void throws_iae_if_top_is_negative() {
      assertThatThrownBy(() -> excel("Margin/no_margin.xlsx").top(-1))
          .isInstanceOf(IllegalArgumentException.class)
          .hasMessage("top must be greater than or equal to 0");
    }

    @Test
    void throws_iae_if_margin_between_header_and_data_is_negative() {
      assertThatThrownBy(() -> excel("Margin/no_margin.xlsx").skipAfterHeader(-1))
          .isInstanceOf(IllegalArgumentException.class)
          .hasMessage("skipAfterHeader must be greater than or equal to 0");
    }
  }

  @Nested
  @SuppressWarnings("ConstantConditions")
  class WithDefaultValue {

    @BeforeEach
    void setUp() {
      var table_11 = "create table if not exists table_11 ("
          + "id integer primary key,"
          + "name varchar(100)"
          + ")";
      var table_12 = "create table if not exists table_12 ("
          + "id integer primary key,"
          + "name varchar(100)"
          + ")";
      var table_13 = "create table if not exists table_13 ("
          + "id integer primary key,"
          + "name varchar(100)"
          + ")";
      new DbSetup(destination, sequenceOf(sql(table_11, table_12, table_13),
          truncate("table_11", "table_12", "table_13"))).launch();
    }

    @Test
    void specify_default_value() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("WithDefaultValue/with_default_value.xlsx")
          .withDefaultValue("table_11", "name", "DEFAULT_1")
          .withDefaultValue("table_13", "name", "DEFAULT_2").build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(5)
          .changeOfCreationOnTable("table_11")
          .rowAtEndPoint()
          .value("id").isEqualTo(10)
          .value("name").isEqualTo("DEFAULT_1")
          .changeOfCreationOnTable("table_11")
          .rowAtEndPoint()
          .value("id").isEqualTo(11)
          .value("name").isEqualTo("DEFAULT_1")
          .changeOfCreationOnTable("table_12")
          .rowAtEndPoint()
          .value("id").isEqualTo(20)
          .value("name").isEqualTo("Bob")
          .changeOfCreationOnTable("table_13")
          .rowAtEndPoint()
          .value("id").isEqualTo(30)
          .value("name").isEqualTo("DEFAULT_2")
          .changeOfCreationOnTable("table_13")
          .rowAtEndPoint()
          .value("id").isEqualTo(31)
          .value("name").isEqualTo("DEFAULT_2");
    }

    @Test
    void throws_npe_if_table_is_null() {
      assertThatThrownBy(() -> excel("WithDefaultValue/with_default_value.xlsx")
          .withDefaultValue(null, "name", "DEFAULT_1"))
          .isInstanceOf(NullPointerException.class)
          .hasMessage("table must not be null");
    }

    @Test
    void throws_npe_if_column_is_null() {
      assertThatThrownBy(() -> excel("WithDefaultValue/with_default_value.xlsx")
          .withDefaultValue("table_11", null, "DEFAULT_1"))
          .isInstanceOf(NullPointerException.class)
          .hasMessage("column must not be null");
    }
  }

  @Nested
  @SuppressWarnings("ConstantConditions")
  class WithGeneratedValue {

    @BeforeEach
    void setUp() {
      var table_11 = "create table if not exists table_11 ("
          + "id integer primary key,"
          + "name varchar(100)"
          + ")";
      var table_12 = "create table if not exists table_12 ("
          + "id integer primary key,"
          + "name varchar(100)"
          + ")";
      var table_13 = "create table if not exists table_13 ("
          + "id integer primary key,"
          + "name varchar(100)"
          + ")";
      new DbSetup(destination, sequenceOf(sql(table_11, table_12, table_13),
          truncate("table_11", "table_12", "table_13"))).launch();
    }

    @Test
    void specify_value_generator() {
      var changes = connection.changes().build().setStartPointNow();
      var operation = excel("WithGeneratedValue/with_generated_value.xlsx")
          .withGeneratedValue("table_11", "id", ValueGenerators.sequence().startingAt(100))
          .withGeneratedValue("table_13", "id", ValueGenerators.sequence().startingAt(300))
          .build();
      new DbSetup(destination, operation).launch();
      assertThat(changes.setEndPointNow())
          .hasNumberOfChanges(5)
          .changeOfCreationOnTable("table_11")
          .rowAtEndPoint()
          .value("id").isEqualTo(100)
          .value("name").isEqualTo("Alice")
          .changeOfCreationOnTable("table_11")
          .rowAtEndPoint()
          .value("id").isEqualTo(101)
          .value("name").isEqualTo("Bob")
          .changeOfCreationOnTable("table_12")
          .rowAtEndPoint()
          .value("id").isEqualTo(2)
          .value("name").isEqualTo("Charlie")
          .changeOfCreationOnTable("table_13")
          .rowAtEndPoint()
          .value("id").isEqualTo(300)
          .value("name").isEqualTo("Dave")
          .changeOfCreationOnTable("table_13")
          .rowAtEndPoint()
          .value("id").isEqualTo(301)
          .value("name").isEqualTo("Erin");
    }

    @Test
    void throws_npe_if_table_is_null() {
      assertThatThrownBy(() -> excel("WithGeneratedValue/with_generated_value.xlsx")
          .withGeneratedValue(null, "name", ValueGenerators.sequence()))
          .isInstanceOf(NullPointerException.class)
          .hasMessage("table must not be null");
    }

    @Test
    void throws_npe_if_column_is_null() {
      assertThatThrownBy(() -> excel("WithGeneratedValue/with_generated_value.xlsx")
          .withGeneratedValue("table_11", null, ValueGenerators.sequence()))
          .isInstanceOf(NullPointerException.class)
          .hasMessage("column must not be null");
    }

    @Test
    void throws_npe_if_generator_is_null() {
      assertThatThrownBy(() -> excel("WithGeneratedValue/with_generated_value.xlsx")
          .withGeneratedValue("table_11", "name", null))
          .isInstanceOf(NullPointerException.class)
          .hasMessage("valueGenerator must not be null");
    }
  }
}
