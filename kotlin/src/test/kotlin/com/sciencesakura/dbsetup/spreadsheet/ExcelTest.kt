// SPDX-License-Identifier: MIT

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
    val changes = connection.changes().build().setStartPointNow()
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
