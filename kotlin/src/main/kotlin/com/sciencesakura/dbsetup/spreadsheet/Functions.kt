// SPDX-License-Identifier: MIT

package com.sciencesakura.dbsetup.spreadsheet

import com.ninja_squad.dbsetup_kotlin.DbSetupBuilder

/**
 * Creates an Excel import operation.
 *
 * @param location the `/`-separated path from classpath root to the Excel file
 * @throws IllegalArgumentException if the Excel file is not found
 */
fun DbSetupBuilder.excel(location: String) {
  execute(Import.excel(location).build())
}

/**
 * Creates an Excel import operation.
 *
 * @param location  the `/`-separated path from classpath root to the Excel file
 * @param configure A lambda to configure the import operation
 * @throws IllegalArgumentException if the Excel file is not found
 */
fun DbSetupBuilder.excel(
  location: String,
  configure: Import.Builder.() -> Unit,
) {
  val excelBuilder = Import.excel(location)
  excelBuilder.configure()
  execute(excelBuilder.build())
}
