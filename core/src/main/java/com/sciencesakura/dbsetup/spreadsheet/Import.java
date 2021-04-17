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

import com.ninja_squad.dbsetup.DbSetupRuntimeException;
import com.ninja_squad.dbsetup.bind.BinderConfiguration;
import com.ninja_squad.dbsetup.generator.ValueGenerator;
import com.ninja_squad.dbsetup.operation.CompositeOperation;
import com.ninja_squad.dbsetup.operation.Insert;
import com.ninja_squad.dbsetup.operation.Operation;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jetbrains.annotations.NotNull;

import java.io.IOException;
import java.net.URL;
import java.sql.Connection;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import static com.sciencesakura.dbsetup.spreadsheet.Cells.a1;
import static com.sciencesakura.dbsetup.spreadsheet.Cells.valueForData;
import static com.sciencesakura.dbsetup.spreadsheet.Cells.valueForHeader;
import static java.util.Objects.requireNonNull;

/**
 * An Operation which imports the Microsoft Excel file into the tables.
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
 *
 * @author sciencesakura
 */
public final class Import implements Operation {

    /**
     * Create a new {@code Import.Builder} instance.
     * <p>
     * The specified location string must be the relative path string from classpath root.
     * </p>
     *
     * @param location the location of the source file that is the relative path from classpath root
     * @return the new {@code Import.Builder} instance
     * @throws IllegalArgumentException if the source file was not found
     */
    @NotNull
    public static Builder excel(@NotNull String location) {
        return new Builder(location);
    }

    private static List<Operation> operations(Builder builder) {
        boolean enableExclude = builder.exclude != null;
        try (Workbook workbook = WorkbookFactory.create(builder.location.openStream())) {
            List<Operation> operations = new ArrayList<>(workbook.getNumberOfSheets());
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                if (workbook.isSheetHidden(i) || workbook.isSheetVeryHidden(i)) continue;
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();
                if (enableExclude && builder.exclude.matcher(sheetName).matches()) continue;
                int rowIndex = builder.top;
                Row row = sheet.getRow(rowIndex++);
                if (row == null) {
                    throw new DbSetupRuntimeException("header row not found: " + sheetName);
                }
                int width = row.getLastCellNum() - builder.left;
                if (width <= 0) {
                    throw new DbSetupRuntimeException("header row not found: " + sheetName);
                }
                Insert.Builder ib = Insert.into(sheetName);
                ib.columns(columns(row, builder.left, width, evaluator));
                Map<String, ValueGenerator<?>> valueGenerators = builder.valueGenerators.get(sheetName);
                if (valueGenerators != null) {
                    valueGenerators.forEach(ib::withGeneratedValue);
                }
                Map<String, ?> defaultValues = builder.defaultValues.get(sheetName);
                if (defaultValues != null) {
                    defaultValues.forEach(ib::withDefaultValue);
                }
                while ((row = sheet.getRow(rowIndex++)) != null) {
                    ib.values(values(row, builder.left, width, evaluator));
                }
                operations.add(ib.build());
            }
            return operations;
        } catch (IOException e) {
            throw new DbSetupRuntimeException("failed to open " + builder.location, e);
        }
    }

    private static String[] columns(Row row, int left, int width, FormulaEvaluator evaluator) {
        String[] columns = new String[width];
        for (int i = 0; i < width; i++) {
            int c = left + i;
            Cell cell = row.getCell(c);
            if (cell == null) {
                throw new DbSetupRuntimeException("header cell not found: " + a1(row.getSheet(), row.getRowNum(), c));
            }
            columns[i] = valueForHeader(cell, evaluator);
        }
        return columns;
    }

    private static Object[] values(Row row, int left, int width, FormulaEvaluator evaluator) {
        Object[] values = new Object[width];
        for (int i = 0; i < width; i++) {
            int c = left + i;
            Cell cell = row.getCell(c);
            values[i] = cell == null ? null : valueForData(cell, evaluator);
        }
        return values;
    }

    private final Operation internalOperation;

    private Import(Builder builder) {
        internalOperation = CompositeOperation.sequenceOf(operations(builder));
    }

    @Override
    public void execute(Connection connection, BinderConfiguration configuration) throws SQLException {
        internalOperation.execute(connection, configuration);
    }

    /**
     * A builder to create the {@code Import} instance.
     *
     * @author sciencesakura
     */
    public static final class Builder {

        private final Map<String, Map<String, Object>> defaultValues = new HashMap<>();

        private final Map<String, Map<String, ValueGenerator<?>>> valueGenerators = new HashMap<>();

        private final URL location;

        private Pattern exclude;

        private int left;

        private int top;

        private Builder(String location) {
            requireNonNull(location, "location must not be null");
            this.location = getClass().getClassLoader().getResource(location);
            if (this.location == null) {
                throw new IllegalArgumentException(location + " not found");
            }
        }

        /**
         * Build a new {@code Import} instance.
         *
         * @return the new {@code Import} instance
         */
        @NotNull
        public Import build() {
            return new Import(this);
        }

        /**
         * Specifies a pattern of the worksheet name to be excluded from the importing.
         *
         * @param exclude the regular expression pattern of the worksheet name to be excluded
         * @return the reference to this object
         */
        public Builder exclude(@NotNull String exclude) {
            this.exclude = Pattern.compile(requireNonNull(exclude, "exclude must not be null"));
            return this;
        }

        /**
         * Specifies a pattern of the worksheet name to be excluded from the importing.
         *
         * @param exclude the regular expression pattern of the worksheet name to be excluded
         * @return the reference to this object
         */
        public Builder exclude(@NotNull Pattern exclude) {
            this.exclude = requireNonNull(exclude, "exclude must not be null");
            return this;
        }

        /**
         * Specifies a start column index of the worksheet to read data.
         * <p>
         * By default 0 is used.
         * </p>
         *
         * @param left the 0-based column index, must be non-negative
         * @return the reference to this object
         */
        public Builder left(int left) {
            if (left < 0) {
                throw new IllegalArgumentException("left must be greater than or equal to 0");
            }
            this.left = left;
            return this;
        }

        /**
         * Specifies a start row index of the worksheet to read data.
         * <p>
         * By default 0 is used.
         * </p>
         *
         * @param top the 0-based row index, must be non-negative
         * @return the reference to this object
         */
        public Builder top(int top) {
            if (top < 0) {
                throw new IllegalArgumentException("top must be greater than or equal to 0");
            }
            this.top = top;
            return this;
        }

        /**
         * Specifies a default value for the given pair of table and column.
         *
         * @param table  the table name
         * @param column the column name
         * @param value  the default value
         * @return the reference to this object
         */
        public Builder withDefaultValue(@NotNull String table, @NotNull String column,
                                        Object value) {
            requireNonNull(table, "table must not be null");
            requireNonNull(column, "column must not be null");
            defaultValues.computeIfAbsent(table, k -> new LinkedHashMap<>()).put(column, value);
            return this;
        }

        /**
         * Specifies a value generator for the given pair of table and column.
         *
         * @param table          the table name
         * @param column         the column name
         * @param valueGenerator the generator
         * @return the reference to this object
         */
        public Builder withGeneratedValue(@NotNull String table, @NotNull String column,
                                          @NotNull ValueGenerator<?> valueGenerator) {
            requireNonNull(table, "table must not be null");
            requireNonNull(column, "column must not be null");
            requireNonNull(valueGenerator, "valueGenerator must not be null");
            valueGenerators.computeIfAbsent(table, k -> new LinkedHashMap<>()).put(column, valueGenerator);
            return this;
        }
    }
}
