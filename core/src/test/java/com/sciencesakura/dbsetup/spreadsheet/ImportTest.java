/*
 * MIT License
 *
 * Copyright (c) 2019 sciencesakura
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */
package com.sciencesakura.dbsetup.spreadsheet;

import static com.ninja_squad.dbsetup.Operations.sql;
import static com.sciencesakura.dbsetup.spreadsheet.Import.excel;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.assertj.db.api.Assertions.assertThat;

import com.ninja_squad.dbsetup.DbSetup;
import com.ninja_squad.dbsetup.DbSetupTracker;
import com.ninja_squad.dbsetup.destination.Destination;
import com.ninja_squad.dbsetup.destination.DriverManagerDestination;
import com.ninja_squad.dbsetup.generator.ValueGenerator;
import com.ninja_squad.dbsetup.generator.ValueGenerators;
import com.ninja_squad.dbsetup.operation.Operation;
import java.util.regex.Pattern;
import org.assertj.db.type.Changes;
import org.assertj.db.type.Source;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

class ImportTest {

    private static final String url = "jdbc:h2:mem:test;DB_CLOSE_DELAY=-1";
    private static final String username = "sa";
    private static final Source source = new Source(url, username, null);
    private static final Destination destination = new DriverManagerDestination(url, username, null);
    private static final DbSetupTracker dbSetupTracker = new DbSetupTracker();
    private static final Operation setUpQueries = sql(
        "create table if not exists table_1 (" +
            "a  integer primary key," +
            "b  bigint," +
            "c  decimal(7, 3)," +
            "d  date," +
            "e  time," +
            "f  timestamp," +
            "g  varchar(6)," +
            "h  boolean," +
            "i  varchar(6)," +
            "dv integer," +
            "gv integer" +
            ")",
        "create table if not exists table_2 (" +
            "a  integer primary key," +
            "b  varchar(6)," +
            "dv char(1)," +
            "gv char(2)" +
            ")",
        "truncate table table_1",
        "truncate table table_2"
    );

    @BeforeEach
    void setUp() {
        dbSetupTracker.launchIfNecessary(new DbSetup(destination, setUpQueries));
    }

    @Test
    void import_workbook_with_default_settings() {
        Changes changes = new Changes(source).setStartPointNow();
        Operation operation = excel("testdata_1.xlsx").build();
        new DbSetup(destination, operation).launch();
        assertThat(changes.setEndPointNow())
            .hasNumberOfChanges(4)
            .changeOnTableWithPks("table_1", 100)
            .isCreation()
            .rowAtEndPoint()
            .value("b").isEqualTo(10000000000L)
            .value("c").isEqualTo(0.5)
            .value("d").isEqualTo("2019-01-01")
            .value("e").isEqualTo("20:30:01")
            .value("f").isEqualTo("2019-02-01T12:34:56")
            .value("g").isEqualTo("甲")
            .value("h").isTrue()
            .value("i").isNull()
            .value("dv").isNull()
            .value("gv").isNull()
            .changeOnTableWithPks("table_1", 200)
            .isCreation()
            .rowAtEndPoint()
            .value("a").isEqualTo(200)
            .value("c").isEqualTo(0.25)
            .value("d").isEqualTo("2019-01-02")
            .value("e").isEqualTo("20:30:02")
            .value("f").isEqualTo("2019-02-02T12:34:56")
            .value("g").isEqualTo("乙")
            .value("h").isFalse()
            .value("i").isNull()
            .value("dv").isNull()
            .value("gv").isNull()
            .changeOnTableWithPks("table_2", 1)
            .isCreation()
            .rowAtEndPoint()
            .value("b").isEqualTo("b1")
            .value("dv").isNull()
            .value("gv").isNull()
            .changeOnTableWithPks("table_2", 2)
            .isCreation()
            .rowAtEndPoint()
            .value("b").isEqualTo("b2")
            .value("dv").isNull()
            .value("gv").isNull();
    }

    @Test
    void import_workbook_with_dv_and_gv() {
        Changes changes = new Changes(source).setStartPointNow();
        Operation operation = excel("testdata_2.xlsx")
            .exclude("table_2_margin", "table_2_error", "table_2_empty")
            .tableMapper(sheet -> sheet.replaceFirst("_\\D+$", ""))
            .withDefaultValue("table_1", "dv", 10)
            .withDefaultValue("table_2", "dv", "X")
            .withGeneratedValue("table_1", "gv", ValueGenerators.sequence())
            .withGeneratedValue("table_2", "gv", ValueGenerators.stringSequence("Y"))
            .build();
        new DbSetup(destination, operation).launch();
        assertThat(changes.setEndPointNow())
            .hasNumberOfChanges(4)
            .changeOnTableWithPks("table_1", 100)
            .isCreation()
            .rowAtEndPoint()
            .value("b").isEqualTo(10000000000L)
            .value("c").isEqualTo(0.5)
            .value("d").isEqualTo("2019-01-01")
            .value("e").isEqualTo("20:30:01")
            .value("f").isEqualTo("2019-02-01T12:34:56")
            .value("g").isEqualTo("甲")
            .value("h").isTrue()
            .value("i").isNull()
            .value("dv").isEqualTo(10)
            .value("gv").isEqualTo(1)
            .changeOnTableWithPks("table_1", 200)
            .isCreation()
            .rowAtEndPoint()
            .value("a").isEqualTo(200)
            .value("c").isEqualTo(0.25)
            .value("d").isEqualTo("2019-01-02")
            .value("e").isEqualTo("20:30:02")
            .value("f").isEqualTo("2019-02-02T12:34:56")
            .value("g").isEqualTo("乙")
            .value("h").isFalse()
            .value("i").isNull()
            .value("dv").isEqualTo(10)
            .value("gv").isEqualTo(2)
            .changeOnTableWithPks("table_2", 1)
            .isCreation()
            .rowAtEndPoint()
            .value("b").isEqualTo("b1")
            .value("dv").isEqualTo("X")
            .value("gv").isEqualTo("Y1")
            .changeOnTableWithPks("table_2", 2)
            .isCreation()
            .rowAtEndPoint()
            .value("b").isEqualTo("b2")
            .value("dv").isEqualTo("X")
            .value("gv").isEqualTo("Y2");
    }

    @Test
    void import_workbook_that_has_margin() {
        Changes changes = new Changes(source).setStartPointNow();
        Operation operation = excel("testdata_2.xlsx")
            .include("table_2.+")
            .exclude("table_2_(formula|error|empty)")
            .tableMapper(sheet -> sheet.replaceFirst("_\\D+$", ""))
            .margin(1, 2)
            .build();
        new DbSetup(destination, operation).launch();
        assertThat(changes.setEndPointNow())
            .hasNumberOfChanges(2)
            .changeOnTableWithPks("table_2", 1)
            .isCreation()
            .rowAtEndPoint()
            .value("b").isEqualTo("b1")
            .changeOnTableWithPks("table_2", 2)
            .isCreation()
            .rowAtEndPoint()
            .value("b").isEqualTo("b2");
    }

    @Test
    void import_workbook_that_contains_error() {
        Operation operation = excel("testdata_2.xlsx")
            .exclude("table_2_(margin|empty)")
            .tableMapper(sheet -> sheet.replaceFirst("_\\D+$", ""))
            .build();
        assertThatThrownBy(() -> new DbSetup(destination, operation).launch())
                .hasMessage("error value contained: table_2_error!B2");
        dbSetupTracker.skipNextLaunch();
    }

    @Test
    void import_workbook_that_contains_empty_sheet() {
        Operation operation = excel("testdata_2.xlsx")
            .exclude("table_2_(margin|error)")
            .tableMapper(sheet -> sheet.replaceFirst("_\\D+$", ""))
            .build();
        assertThatThrownBy(() -> new DbSetup(destination, operation).launch())
                .hasMessage("header row not found: table_2_empty[0]");
        dbSetupTracker.skipNextLaunch();
    }

    static class IllegalArgument {

        @Test
        void file_not_found() {
            assertThatThrownBy(() -> excel("xxx"))
                    .hasMessage("xxx not found");
        }

        @Test
        void location_is_null() {
            String location = null;
            assertThatThrownBy(() -> excel(location))
                    .hasMessage("location must not be null");
        }

        @Test
        void string_include_is_null() {
            assertThatThrownBy(() -> excel("testdata_1.xlsx")
                .include((String[]) null))
                .hasMessage("regex must not be null");
        }

        @Test
        void string_include_contains_null() {
            assertThatThrownBy(() -> excel("testdata_1.xlsx")
                .include((String) null))
                .hasMessage("regex must not contain null");
        }

        @Test
        void pattern_include_is_null() {
            assertThatThrownBy(() -> excel("testdata_1.xlsx")
                .include((Pattern[]) null))
                .hasMessage("patterns must not be null");
        }

        @Test
        void pattern_include_contains_null() {
            assertThatThrownBy(() -> excel("testdata_1.xlsx")
                .include((Pattern) null))
                .hasMessage("patterns must not contain null");
        }

        @Test
        void string_exclude_is_null() {
            assertThatThrownBy(() -> excel("testdata_1.xlsx")
                .exclude((String[]) null))
                .hasMessage("regex must not be null");
        }

        @Test
        void string_exclude_contains_null() {
            assertThatThrownBy(() -> excel("testdata_1.xlsx")
                .exclude((String) null))
                .hasMessage("regex must not contain null");
        }

        @Test
        void pattern_exclude_is_null() {
            assertThatThrownBy(() -> excel("testdata_1.xlsx")
                .exclude((Pattern[]) null))
                .hasMessage("patterns must not be null");
        }

        @Test
        void pattern_exclude_contains_null() {
            assertThatThrownBy(() -> excel("testdata_1.xlsx")
                .exclude((Pattern) null))
                .hasMessage("patterns must not contain null");
        }

        @Test
        void left_is_negative() {
            int left = -1;
            assertThatThrownBy(() -> excel("testdata_1.xlsx").left(left))
                    .hasMessage("left must be greater than or equal to 0");
        }

        @Test
        void top_is_negative() {
            int top = -1;
            assertThatThrownBy(() -> excel("testdata_1.xlsx").top(top))
                    .hasMessage("top must be greater than or equal to 0");
        }

        @Test
        void default_value_table_is_null() {
            String table = null;
            String column = "column";
            Object value = new Object();
            assertThatThrownBy(() -> excel("testdata_1.xlsx")
                    .withDefaultValue(table, column, value))
                    .hasMessage("table must not be null");
        }

        @Test
        void default_value_column_is_null() {
            String table = "table";
            String column = null;
            Object value = new Object();
            assertThatThrownBy(() -> excel("testdata_1.xlsx")
                    .withDefaultValue(table, column, value))
                    .hasMessage("column must not be null");
        }

        @Test
        void value_generator_table_is_null() {
            String table = null;
            String column = "column";
            ValueGenerator<?> valueGenerator = ValueGenerators.sequence();
            assertThatThrownBy(() -> excel("testdata_1.xlsx")
                    .withGeneratedValue(table, column, valueGenerator))
                    .hasMessage("table must not be null");
        }

        @Test
        void value_generator_column_is_null() {
            String table = "table";
            String column = null;
            ValueGenerator<?> valueGenerator = ValueGenerators.sequence();
            assertThatThrownBy(() -> excel("testdata_1.xlsx")
                    .withGeneratedValue(table, column, valueGenerator))
                    .hasMessage("column must not be null");
        }

        @Test
        void value_generator_generator_is_null() {
            String table = "table";
            String column = "column";
            ValueGenerator<?> valueGenerator = null;
            assertThatThrownBy(() -> excel("testdata_1.xlsx")
                    .withGeneratedValue(table, column, valueGenerator))
                    .hasMessage("valueGenerator must not be null");
        }
    }
}
