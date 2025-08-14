// SPDX-License-Identifier: MIT

package com.sciencesakura.dbsetup.spreadsheet

import com.ninja_squad.dbsetup_kotlin.DbSetupBuilder

/**
 * Add an Excel import operation to the `DbSetupBuilder`.
 *
 * @param location the location of the source file that is the relative path from classpath root
 * @throws IllegalArgumentException if the source file was not found
 */
fun DbSetupBuilder.excel(location: String) {
    execute(Import.excel(location).build())
}

/**
 * Add an Excel import operation to the `DbSetupBuilder`.
 *
 * @param location  the location of the source file that is the relative path from classpath root
 * @param configure the function used to configure the Excel import
 * @throws IllegalArgumentException if the source file was not found
 */
fun DbSetupBuilder.excel(
    location: String,
    configure: Import.Builder.() -> Unit,
) {
    val excelBuilder = Import.excel(location)
    excelBuilder.configure()
    execute(excelBuilder.build())
}
