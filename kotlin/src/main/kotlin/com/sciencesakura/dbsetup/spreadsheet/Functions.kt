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
