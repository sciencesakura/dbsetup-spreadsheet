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
import com.ninja_squad.dbsetup_kotlin.dbSetup
import org.assertj.db.api.Assertions.assertThat
import org.assertj.db.type.AssertDbConnectionFactory
import kotlin.test.BeforeTest
import kotlin.test.Test

class ExcelTest {
    val url = "jdbc:h2:mem:test;DB_CLOSE_DELAY=-1"

    val username = "sa"

    val connection = AssertDbConnectionFactory.of(url, username, null).create()

    val destination = DriverManagerDestination.with(url, username, null)

    @BeforeTest
    fun setUp() {
        val table11 =
            """
            create table if not exists table_11 (
              id integer primary key,
              name varchar(100)
            )
            """.trimIndent()
        val table12 =
            """
            create table if not exists table_12 (
              id integer primary key,
              name varchar(100)
            )
            """.trimIndent()
        dbSetup(destination) {
            sql(table11, table12)
            truncate("table_11", "table_12")
        }.launch()
    }

    @Test
    fun import_excel() {
        val changes = connection.changes().build().setStartPointNow()
        dbSetup(destination) {
            excel("kt_test.xlsx")
        }.launch()
        assertThat(changes.setEndPointNow()).hasNumberOfChanges(2).changeOfCreationOnTable("table_11")
            .rowAtEndPoint().value("id").isEqualTo(1).value("name").isEqualTo("Alice")
            .changeOfCreationOnTable("table_12").rowAtEndPoint().value("id").isEqualTo(2).value("name")
            .isEqualTo("Bob")
    }

    @Test
    fun import_excel_with_configure() {
      val changes = connection.changes().build().setStartPointNow()
        dbSetup(destination) {
            excel("kt_test.xlsx") {
                exclude("table_12")
            }
        }.launch()
        assertThat(changes.setEndPointNow()).hasNumberOfChanges(1).changeOfCreationOnTable("table_11")
            .rowAtEndPoint().value("id").isEqualTo(1).value("name").isEqualTo("Alice")
    }
}
