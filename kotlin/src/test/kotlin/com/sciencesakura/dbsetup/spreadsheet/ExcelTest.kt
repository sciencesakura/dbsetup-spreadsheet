// SPDX-License-Identifier: MIT

package com.sciencesakura.dbsetup.spreadsheet

import com.ninja_squad.dbsetup.destination.Destination
import com.ninja_squad.dbsetup.destination.DriverManagerDestination
import com.ninja_squad.dbsetup_kotlin.dbSetup
import org.assertj.db.api.Assertions.assertThat
import org.assertj.db.type.AssertDbConnectionFactory
import org.assertj.db.type.Changes
import kotlin.test.BeforeTest
import kotlin.test.Test

class ExcelTest {
  lateinit var destination: Destination

  lateinit var changes: Changes

  @BeforeTest
  fun setUp() {
    val url = "jdbc:h2:mem:test;DB_CLOSE_DELAY=-1"
    val username = "sa"
    val connection = AssertDbConnectionFactory.of(url, username, null).create()
    destination = DriverManagerDestination.with(url, username, null)
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
    changes = connection.changes().build()
  }

  @Test
  fun import_excel() {
    changes.setStartPointNow()
    dbSetup(destination) {
      excel("kt_test.xlsx")
    }.launch()
    @Suppress("ktlint:standard:chain-method-continuation")
    assertThat(changes.setEndPointNow())
      .hasNumberOfChanges(2)
      .changeOfCreationOnTable("table_11")
      .rowAtEndPoint()
      .value("id").isEqualTo(1)
      .value("name").isEqualTo("Alice")
      .changeOfCreationOnTable("table_12")
      .rowAtEndPoint()
      .value("id").isEqualTo(2)
      .value("name").isEqualTo("Bob")
  }

  @Test
  fun import_excel_with_configure() {
    changes.setStartPointNow()
    dbSetup(destination) {
      excel("kt_test.xlsx") {
        exclude("table_12")
      }
    }.launch()
    @Suppress("ktlint:standard:chain-method-continuation")
    assertThat(changes.setEndPointNow())
      .hasNumberOfChanges(1)
      .changeOfCreationOnTable("table_11")
      .rowAtEndPoint()
      .value("id").isEqualTo(1)
      .value("name").isEqualTo("Alice")
  }
}
