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
/**
 * A package containing classes to import data from Microsoft Excel files.
 * <p>
 * Usage:
 * </p>
 * <p>
 * When there are tables below:
 * </p>
 * <pre style="background-color: #f6f8fa;"><code>
 * create table country (
 *   id    integer       not null,
 *   code  char(3)       not null,
 *   name  varchar(256)  not null,
 *   primary key (id),
 *   unique (code)
 * );
 *
 * create table customer (
 *   id      integer       not null,
 *   name    varchar(256)  not null,
 *   country integer       not null,
 *   primary key (id),
 *   foreign key (country) references country (id)
 * );
 * </code></pre>
 * <p>
 * Create An Excel file with one worksheet per table, and name those worksheets
 * the same name as the tables. If there is a dependency between the tables,
 * put the worksheet of dependent table early.
 * </p>
 * <table class="striped">
 *   <caption>country sheet</caption>
 *   <tbody>
 *     <tr>
 *       <th>id</th>
 *       <th>code</th>
 *       <th>name</th>
 *     </tr>
 *     <tr>
 *       <td>1</td>
 *       <td>GBR</td>
 *       <td>United Kingdom</td>
 *     </tr>
 *     <tr>
 *       <td>2</td>
 *       <td>HKG</td>
 *       <td>Hong Kong</td>
 *     </tr>
 *     <tr>
 *       <td>3</td>
 *       <td>JPN</td>
 *       <td>Japan</td>
 *     </tr>
 *   </tbody>
 * </table>
 * <table class="striped">
 *   <caption>customer sheet</caption>
 *   <tbody>
 *     <tr>
 *       <th>id</th>
 *       <th>name</th>
 *       <th>country</th>
 *     </tr>
 *     <tr>
 *       <td>1</td>
 *       <td>Eriol</td>
 *       <td>1</td>
 *     </tr>
 *     <tr>
 *       <td>2</td>
 *       <td>Sakura</td>
 *       <td>3</td>
 *     </tr>
 *     <tr>
 *       <td>3</td>
 *       <td>Xiaolang</td>
 *       <td>2</td>
 *     </tr>
 *   </tbody>
 * </table>
 * <p>
 * Put the prepared Excel file on the classpath, and write code like below:
 * </p>
 * <pre style="background-color: #f6f8fa;"><code>
 * import static com.sciencesakura.dbsetup.spreadsheet.Import.excel;
 *
 * Operation operation = excel("testdata.xlsx").build();
 * DbSetup dbSetup = new DbSetup(destination, operation);
 * dbSetup.launch();
 * </code></pre>
 */
package com.sciencesakura.dbsetup.spreadsheet;
