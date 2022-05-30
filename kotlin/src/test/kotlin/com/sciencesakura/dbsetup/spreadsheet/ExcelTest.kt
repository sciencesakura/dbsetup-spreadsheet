/*
 * MIT License
 *
 * Copyright (c) 2020 sciencesakura
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
package com.sciencesakura.dbsetup.spreadsheet

import com.ninja_squad.dbsetup.destination.DriverManagerDestination
import com.ninja_squad.dbsetup.generator.ValueGenerators
import com.ninja_squad.dbsetup_kotlin.dbSetup
import org.assertj.db.api.Assertions.assertThat
import org.assertj.db.type.Changes
import org.assertj.db.type.Source
import org.junit.jupiter.api.BeforeEach
import org.junit.jupiter.api.Test

private const val url = "jdbc:h2:mem:test;DB_CLOSE_DELAY=-1"
private const val username = "sa"
private val source = Source(url, username, null)
private val destination = DriverManagerDestination(url, username, null)
private val setUpQueries = arrayOf(
    """
    create table if not exists table_1 (
      a integer primary key,
      b integer,
      c integer
    )
    """,
    "truncate table table_1"
)

class ExcelTest {

    @BeforeEach
    fun setUp() {
        dbSetup(destination) {
            sql(*setUpQueries)
        }.launch()
    }

    @Test
    fun import_with_default_settings() {
        val changes = Changes(source).setStartPointNow()
        dbSetup(destination) {
            excel("testdata.xlsx")
        }.launch()
        assertThat(changes.setEndPointNow())
            .hasNumberOfChanges(2)
            .changeOfCreation()
            .rowAtEndPoint()
            .value("a").isEqualTo(10)
            .value("b").isEqualTo(100)
            .value("c").isNull
            .changeOfCreation()
            .rowAtEndPoint()
            .value("a").isEqualTo(20)
            .value("b").isEqualTo(200)
            .value("c").isNull
    }

    @Test
    fun import_with_customized_settings() {
        val changes = Changes(source).setStartPointNow()
        dbSetup(destination) {
            excel("testdata.xlsx") {
                withGeneratedValue("table_1", "c", ValueGenerators.sequence())
            }
        }.launch()
        assertThat(changes.setEndPointNow())
            .hasNumberOfChanges(2)
            .changeOfCreation()
            .rowAtEndPoint()
            .value("a").isEqualTo(10)
            .value("b").isEqualTo(100)
            .value("c").isEqualTo(1)
            .changeOfCreation()
            .rowAtEndPoint()
            .value("a").isEqualTo(20)
            .value("b").isEqualTo(200)
            .value("c").isEqualTo(2)
    }
}
