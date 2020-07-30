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

import com.ninja_squad.dbsetup.DbSetup;
import com.ninja_squad.dbsetup.destination.Destination;
import com.ninja_squad.dbsetup.destination.DriverManagerDestination;
import com.ninja_squad.dbsetup.generator.ValueGenerator;
import com.ninja_squad.dbsetup.generator.ValueGenerators;
import com.ninja_squad.dbsetup.operation.Operation;
import org.assertj.db.type.Changes;
import org.assertj.db.type.Source;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import static com.ninja_squad.dbsetup.Operations.sql;
import static com.sciencesakura.dbsetup.spreadsheet.Import.excel;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.assertj.db.api.Assertions.assertThat;

class ImportTest {

    private static final String url = "jdbc:h2:mem:test;DB_CLOSE_DELAY=-1";

    private static final String username = "sa";

    private final Destination destination = new DriverManagerDestination(url, username, null);

    private final Source source = new Source(url, username, null);

    @BeforeEach
    void setUp() {
        String[] ddl = {
                "drop table if exists table_1 cascade",
                "create table table_1 (" +
                "  a   integer primary key," +
                "  b   bigint," +
                "  c   decimal(7, 3)," +
                "  d   date," +
                "  e   timestamp," +
                "  f   char(3)," +
                "  g   varchar(6)," +
                "  h   boolean," +
                "  i   varchar(6)" +
                ")",
                "drop table if exists table_2 cascade",
                "create table table_2 (" +
                "  a   integer primary key," +
                "  b   varchar(6)" +
                ")"
        };
        new DbSetup(destination, sql(ddl)).launch();
    }

    @Test
    void single_sheet() {
        Changes changes = new Changes(source).setStartPointNow();
        Operation operation = excel("single_sheet.xlsx").build();
        new DbSetup(destination, operation).launch();
        changes.setEndPointNow();
        assertThat(changes).hasNumberOfChanges(2)
                .changeOfCreationOnTable("table_1")
                .rowAtEndPoint()
                .value("a").isEqualTo(100)
                .value("b").isEqualTo(10000000000L)
                .value("c").isEqualTo(0.5)
                .value("d").isEqualTo("2019-12-01")
                .value("e").isEqualTo("2019-12-01T09:30:01.001000000")
                .value("f").isEqualTo("AAA")
                .value("g").isEqualTo("甲")
                .value("h").isTrue()
                .value("i").isNotNull()
                .changeOfCreationOnTable("table_1")
                .rowAtEndPoint()
                .value("a").isEqualTo(200)
                .value("b").isEqualTo(20000000000L)
                .value("c").isEqualTo(0.25)
                .value("d").isEqualTo("2019-12-02")
                .value("e").isEqualTo("2019-12-02T09:30:02.002000000")
                .value("f").isEqualTo("BBB")
                .value("g").isEqualTo("乙")
                .value("h").isFalse()
                .value("i").isNull();
    }

    @Test
    void multiple_sheet() {
        Changes changes = new Changes(source).setStartPointNow();
        Operation operation = excel("multiple_sheet.xlsx").build();
        new DbSetup(destination, operation).launch();
        changes.setEndPointNow();
        assertThat(changes).hasNumberOfChanges(5)
                .changeOfCreationOnTable("table_1")
                .rowAtEndPoint()
                .value("a").isEqualTo(100)
                .value("b").isEqualTo(10000000000L)
                .value("c").isEqualTo(0.5)
                .value("d").isEqualTo("2019-12-01")
                .value("e").isEqualTo("2019-12-01T09:30:01.001000000")
                .value("f").isEqualTo("AAA")
                .value("g").isEqualTo("甲")
                .value("h").isTrue()
                .value("i").isNotNull()
                .changeOfCreationOnTable("table_1")
                .rowAtEndPoint()
                .value("a").isEqualTo(200)
                .value("b").isEqualTo(20000000000L)
                .value("c").isEqualTo(0.25)
                .value("d").isEqualTo("2019-12-02")
                .value("e").isEqualTo("2019-12-02T09:30:02.002000000")
                .value("f").isEqualTo("BBB")
                .value("g").isEqualTo("乙")
                .value("h").isFalse()
                .value("i").isNull()
                .changeOfCreationOnTable("table_2")
                .rowAtEndPoint()
                .value("a").isEqualTo(101)
                .value("b").isEqualTo("AAA")
                .changeOfCreationOnTable("table_2")
                .rowAtEndPoint()
                .value("a").isEqualTo(201)
                .value("b").isEqualTo("BBB")
                .changeOfCreationOnTable("table_2")
                .rowAtEndPoint()
                .value("a").isEqualTo(301)
                .value("b").isEqualTo("CCC");
    }

    @Test
    void generators() {
        Changes changes = new Changes(source).setStartPointNow();
        Operation operation = excel("generators.xlsx")
                .withGeneratedValue("table_1", "a", ValueGenerators.sequence())
                .withGeneratedValue("table_1", "g", ValueGenerators.stringSequence("G-"))
                .withGeneratedValue("table_2", "a", ValueGenerators.sequence().startingAt(101))
                .build();
        new DbSetup(destination, operation).launch();
        changes.setEndPointNow();
        assertThat(changes).hasNumberOfChanges(5)
                .changeOfCreationOnTable("table_1")
                .rowAtEndPoint()
                .value("a").isEqualTo(1)
                .value("g").isEqualTo("G-1")
                .changeOfCreationOnTable("table_1")
                .rowAtEndPoint()
                .value("a").isEqualTo(2)
                .value("g").isEqualTo("G-2")
                .changeOfCreationOnTable("table_2")
                .rowAtEndPoint()
                .value("a").isEqualTo(101)
                .changeOfCreationOnTable("table_2")
                .rowAtEndPoint()
                .value("a").isEqualTo(102)
                .changeOfCreationOnTable("table_2")
                .rowAtEndPoint()
                .value("a").isEqualTo(103);
    }

    @Test
    void constants() {
        Changes changes = new Changes(source).setStartPointNow();
        Operation operation = excel("constants.xlsx")
                .withDefaultValue("table_1", "g", "G")
                .withDefaultValue("table_1", "i", null)
                .withDefaultValue("table_2", "b", "B")
                .build();
        new DbSetup(destination, operation).launch();
        changes.setEndPointNow();
        assertThat(changes).hasNumberOfChanges(5)
                .changeOfCreationOnTable("table_1")
                .rowAtEndPoint()
                .value("g").isEqualTo("G")
                .value("i").isNull()
                .changeOfCreationOnTable("table_1")
                .rowAtEndPoint()
                .value("g").isEqualTo("G")
                .value("i").isNull()
                .changeOfCreationOnTable("table_2")
                .rowAtEndPoint()
                .value("b").isEqualTo("B")
                .changeOfCreationOnTable("table_2")
                .rowAtEndPoint()
                .value("b").isEqualTo("B")
                .changeOfCreationOnTable("table_2")
                .rowAtEndPoint()
                .value("b").isEqualTo("B");
    }

    @Test
    void has_margin() {
        Changes changes = new Changes(source).setStartPointNow();
        Operation operation = excel("has_margin.xlsx")
                .top(2)
                .left(1)
                .build();
        new DbSetup(destination, operation).launch();
        changes.setEndPointNow();
        assertThat(changes).hasNumberOfChanges(3)
                .changeOfCreationOnTable("table_2")
                .rowAtEndPoint()
                .value("a").isEqualTo(101)
                .value("b").isEqualTo("AAA")
                .changeOfCreationOnTable("table_2")
                .rowAtEndPoint()
                .value("a").isEqualTo(201)
                .value("b").isEqualTo("BBB")
                .changeOfCreationOnTable("table_2")
                .rowAtEndPoint()
                .value("a").isEqualTo(301)
                .value("b").isEqualTo("CCC");
    }

    @Test
    void contains_formula() {
        Changes changes = new Changes(source).setStartPointNow();
        Operation operation = excel("contains_formula.xlsx").build();
        new DbSetup(destination, operation).launch();
        changes.setEndPointNow();
        assertThat(changes).hasNumberOfChanges(1)
                .changeOfCreationOnTable("table_1")
                .rowAtEndPoint()
                .value("a").isEqualTo(10)
                .value("b").isEqualTo(21)
                .value("c").isEqualTo(10.5)
                .value("d").isEqualTo("2019-10-21")
                .value("f").isEqualTo("abc")
                .value("h").isFalse();
    }

    @Test
    void contains_error() {
        Import.Builder ib = excel("contains_error.xlsx");
        assertThatThrownBy(ib::build)
                .hasMessage("error value contained: table_2!B4");
    }

    @Test
    void contains_empty_sheet() {
        Import.Builder ib = excel("contains_empty_sheet.xlsx");
        assertThatThrownBy(ib::build)
                .hasMessage("header not found: empty_sheet");
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
        void left_is_negative() {
            int left = -1;
            assertThatThrownBy(() -> excel("single_sheet.xlsx").left(left))
                    .hasMessage("left must be greater than or equal to 0");
        }

        @Test
        void top_is_negative() {
            int top = -1;
            assertThatThrownBy(() -> excel("single_sheet.xlsx").top(top))
                    .hasMessage("top must be greater than or equal to 0");
        }

        @Test
        void default_value_table_is_null() {
            String table = null;
            String column = "column";
            Object value = new Object();
            assertThatThrownBy(() -> excel("single_sheet.xlsx")
                    .withDefaultValue(table, column, value))
                    .hasMessage("table must not be null");
        }

        @Test
        void default_value_column_is_null() {
            String table = "table";
            String column = null;
            Object value = new Object();
            assertThatThrownBy(() -> excel("single_sheet.xlsx")
                    .withDefaultValue(table, column, value))
                    .hasMessage("column must not be null");
        }

        @Test
        void value_generator_table_is_null() {
            String table = null;
            String column = "column";
            ValueGenerator<?> valueGenerator = ValueGenerators.sequence();
            assertThatThrownBy(() -> excel("single_sheet.xlsx")
                    .withGeneratedValue(table, column, valueGenerator))
                    .hasMessage("table must not be null");
        }

        @Test
        void value_generator_column_is_null() {
            String table = "table";
            String column = null;
            ValueGenerator<?> valueGenerator = ValueGenerators.sequence();
            assertThatThrownBy(() -> excel("single_sheet.xlsx")
                    .withGeneratedValue(table, column, valueGenerator))
                    .hasMessage("column must not be null");
        }

        @Test
        void value_generator_generator_is_null() {
            String table = "table";
            String column = "column";
            ValueGenerator<?> valueGenerator = null;
            assertThatThrownBy(() -> excel("single_sheet.xlsx")
                    .withGeneratedValue(table, column, valueGenerator))
                    .hasMessage("valueGenerator must not be null");
        }
    }

    static class IllegalState {

        @Test
        void build_after_built() {
            Import.Builder ib = excel("single_sheet.xlsx");
            ib.build();
            assertThatThrownBy(ib::build)
                    .isInstanceOf(IllegalStateException.class)
                    .hasMessage("this operation has been built already");
        }

        @Test
        void left_after_built() {
            Import.Builder ib = excel("single_sheet.xlsx");
            ib.build();
            assertThatThrownBy(() -> ib.left(1))
                    .isInstanceOf(IllegalStateException.class)
                    .hasMessage("this operation has been built already");
        }

        @Test
        void top_after_built() {
            Import.Builder ib = excel("single_sheet.xlsx");
            ib.build();
            assertThatThrownBy(() -> ib.top(1))
                    .isInstanceOf(IllegalStateException.class)
                    .hasMessage("this operation has been built already");
        }

        @Test
        void default_value_after_built() {
            Import.Builder ib = excel("single_sheet.xlsx");
            ib.build();
            assertThatThrownBy(() -> ib.withDefaultValue("table", "column", "value"))
                    .isInstanceOf(IllegalStateException.class)
                    .hasMessage("this operation has been built already");
        }

        @Test
        void value_generator_after_built() {
            Import.Builder ib = excel("single_sheet.xlsx");
            ib.build();
            assertThatThrownBy(() -> ib.withGeneratedValue("table", "column", ValueGenerators.sequence()))
                    .isInstanceOf(IllegalStateException.class)
                    .hasMessage("this operation has been built already");
        }
    }
}
