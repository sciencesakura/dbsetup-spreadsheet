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
import com.ninja_squad.dbsetup.destination.DataSourceDestination;
import com.ninja_squad.dbsetup.destination.Destination;
import com.ninja_squad.dbsetup.operation.Operation;
import org.assertj.db.type.Table;
import org.flywaydb.core.Flyway;
import org.flywaydb.core.api.configuration.FluentConfiguration;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import javax.sql.DataSource;

import static com.ninja_squad.dbsetup.Operations.sequenceOf;
import static com.ninja_squad.dbsetup.Operations.truncate;
import static com.sciencesakura.dbsetup.spreadsheet.Import.excel;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.assertj.db.api.Assertions.assertThat;
import static org.assertj.db.type.Table.Order.asc;

class ImportTest {

    private static final Table.Order[] ORDER_BY_A = {asc("a")};

    private static DataSource dataSource;

    private static Destination destination;

    @BeforeAll
    static void setUpClass() {
        String url = "jdbc:h2:mem:ImportTest;DB_CLOSE_DELAY=-1";
        FluentConfiguration conf = Flyway.configure().dataSource(url, "sa", null);
        conf.load().migrate();
        dataSource = conf.getDataSource();
        destination = DataSourceDestination.with(dataSource);
    }

    @Test
    void single_sheet() {
        Operation operation = sequenceOf(
                truncate("table_1"),
                excel("data/single_sheet.xlsx").build());
        DbSetup dbSetup = new DbSetup(destination, operation);
        dbSetup.launch();
        assertThat(new Table(dataSource, "table_1", ORDER_BY_A))
                .row()
                .column("a").value().isEqualTo(100)
                .column("b").value().isEqualTo(10000000000L)
                .column("c").value().isEqualTo(0.5)
                .column("d").value().isEqualTo("2019-12-01")
                .column("e").value().isEqualTo("2019-12-01T09:30:01.001000000")
                .column("f").value().isEqualTo("AAA")
                .column("g").value().isEqualTo("甲")
                .column("h").value().isTrue()
                .column("i").value().isNotNull()
                .row()
                .column("a").value().isEqualTo(200)
                .column("b").value().isEqualTo(20000000000L)
                .column("c").value().isEqualTo(0.25)
                .column("d").value().isEqualTo("2019-12-02")
                .column("e").value().isEqualTo("2019-12-02T09:30:02.002000000")
                .column("f").value().isEqualTo("BBB")
                .column("g").value().isEqualTo("乙")
                .column("h").value().isFalse()
                .column("i").value().isNull();
    }

    @Test
    void multiple_sheet() {
        Operation operation = sequenceOf(
                truncate("table_1", "table_2"),
                excel("data/multiple_sheet.xlsx").build());
        DbSetup dbSetup = new DbSetup(destination, operation);
        dbSetup.launch();
        assertThat(new Table(dataSource, "table_1", ORDER_BY_A))
                .row()
                .column("a").value().isEqualTo(100)
                .column("b").value().isEqualTo(10000000000L)
                .column("c").value().isEqualTo(0.5)
                .column("d").value().isEqualTo("2019-12-01")
                .column("e").value().isEqualTo("2019-12-01T09:30:01.001000000")
                .column("f").value().isEqualTo("AAA")
                .column("g").value().isEqualTo("甲")
                .column("h").value().isTrue()
                .column("i").value().isNotNull()
                .row()
                .column("a").value().isEqualTo(200)
                .column("b").value().isEqualTo(20000000000L)
                .column("c").value().isEqualTo(0.25)
                .column("d").value().isEqualTo("2019-12-02")
                .column("e").value().isEqualTo("2019-12-02T09:30:02.002000000")
                .column("f").value().isEqualTo("BBB")
                .column("g").value().isEqualTo("乙")
                .column("h").value().isFalse()
                .column("i").value().isNull();
        assertThat(new Table(dataSource, "table_2", ORDER_BY_A))
                .row()
                .column("a").value().isEqualTo(101)
                .column("b").value().isEqualTo("AAA")
                .row()
                .column("a").value().isEqualTo(201)
                .column("b").value().isEqualTo("BBB")
                .row()
                .column("a").value().isEqualTo(301)
                .column("b").value().isEqualTo("CCC");
    }

    @Test
    void has_margin() {
        Operation operation = sequenceOf(
                truncate("table_2"),
                excel("data/has_margin.xlsx")
                        .top(2)
                        .left(1)
                        .build());
        DbSetup dbSetup = new DbSetup(destination, operation);
        dbSetup.launch();
        assertThat(new Table(dataSource, "table_2", ORDER_BY_A))
                .row()
                .column("a").value().isEqualTo(101)
                .column("b").value().isEqualTo("AAA")
                .row()
                .column("a").value().isEqualTo(201)
                .column("b").value().isEqualTo("BBB")
                .row()
                .column("a").value().isEqualTo(301)
                .column("b").value().isEqualTo("CCC");
    }

    @Test
    void contains_formula() {
        Operation operation = sequenceOf(
                truncate("table_1"),
                excel("data/contains_formula.xlsx").build());
        DbSetup dbSetup = new DbSetup(destination, operation);
        dbSetup.launch();
        assertThat(new Table(dataSource, "table_1", ORDER_BY_A))
                .row()
                .column("a").value().isEqualTo(10)
                .column("b").value().isEqualTo(21)
                .column("c").value().isEqualTo(10.5)
                .column("d").value().isEqualTo("2019-10-21")
                .column("f").value().isEqualTo("abc")
                .column("h").value().isFalse();
    }

    @Test
    void contains_error() {
        Import.Builder ib = excel("data/contains_error.xlsx");
        assertThatThrownBy(ib::build)
                .hasMessage("error value contained: table_2!B4");
    }

    @Test
    void contains_empty_sheet() {
        Import.Builder ib = excel("data/contains_empty_sheet.xlsx");
        assertThatThrownBy(ib::build)
                .hasMessage("header not found: empty_sheet");
    }

    @Nested
    class IllegalArgument {

        @Test
        void file_not_found() {
            assertThatThrownBy(() -> excel("xxx"))
                    .hasMessage("xxx not found");
        }

        @Test
        void left_is_negative() {
            int left = -1;
            assertThatThrownBy(() -> excel("data/single_sheet.xlsx").left(left))
                    .hasMessage("left must be greater than or equal to 0");
        }

        @Test
        void top_is_negative() {
            int top = -1;
            assertThatThrownBy(() -> excel("data/single_sheet.xlsx").top(top))
                    .hasMessage("top must be greater than or equal to 0");
        }
    }

    @Nested
    class IllegalState {

        @Test
        void build_after_built() {
            Import.Builder ib = excel("data/single_sheet.xlsx");
            ib.build();
            assertThatThrownBy(ib::build)
                    .isInstanceOf(IllegalStateException.class)
                    .hasMessage("this operation has been built already");
        }

        @Test
        void left_after_built() {
            Import.Builder ib = excel("data/single_sheet.xlsx");
            ib.build();
            assertThatThrownBy(() -> ib.left(1))
                    .isInstanceOf(IllegalStateException.class)
                    .hasMessage("this operation has been built already");
        }

        @Test
        void top_after_built() {
            Import.Builder ib = excel("data/single_sheet.xlsx");
            ib.build();
            assertThatThrownBy(() -> ib.top(1))
                    .isInstanceOf(IllegalStateException.class)
                    .hasMessage("this operation has been built already");
        }
    }
}
