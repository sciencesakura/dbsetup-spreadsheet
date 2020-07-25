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

import static com.sciencesakura.dbsetup.spreadsheet.Cells.a1;
import static com.sciencesakura.dbsetup.spreadsheet.Cells.valueForData;
import static com.sciencesakura.dbsetup.spreadsheet.Cells.valueForHeader;
import static java.util.Objects.requireNonNull;

/**
 * An Operation which imports a Microsoft Excel file into the tables.
 *
 * @author sciencesakura
 */
public class Import implements Operation {

    /**
     * Create a new builder.
     * <p>
     * The specified location string must be the relative path string from classpath root.
     * </p>
     *
     * @param location a location of source file that is the relative path from classpath root
     * @return a new Import.Builder
     */
    @NotNull
    public static Builder excel(@NotNull String location) {
        return new Builder(location);
    }

    private static String[] columns(Row row, int left, int width, FormulaEvaluator evaluator) {
        String[] columns = new String[width];
        for (int i = 0; i < width; i++) {
            int c = left + i;
            Cell cell = row.getCell(c);
            if (cell == null)
                throw new DbSetupRuntimeException("header cell not found: " + a1(row.getSheet(), row.getRowNum(), c));
            columns[i] = valueForHeader(cell, evaluator);
        }
        return columns;
    }

    private static List<Operation> operations(Builder builder) {
        int top = builder.top;
        int left = builder.left;
        try (Workbook workbook = WorkbookFactory.create(builder.location.openStream())) {
            List<Operation> operations = new ArrayList<>(workbook.getNumberOfSheets());
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            for (Sheet sheet : workbook) {
                int rowIndex = top;
                Row row = sheet.getRow(rowIndex++);
                if (row == null)
                    throw new DbSetupRuntimeException("header not found: " + sheet.getSheetName());
                int width = row.getLastCellNum() - left;
                if (width <= 0)
                    throw new DbSetupRuntimeException("header not found: " + sheet.getSheetName());
                Insert.Builder ib = Insert.into(sheet.getSheetName());
                ib.columns(columns(row, left, width, evaluator));
                Map<String, ValueGenerator<?>> valueGenerators = builder.valueGenerators.get(sheet.getSheetName());
                if (valueGenerators != null) {
                    valueGenerators.forEach(ib::withGeneratedValue);
                }
                while ((row = sheet.getRow(rowIndex++)) != null) {
                    ib.values(values(row, left, width, evaluator));
                }
                operations.add(ib.build());
            }
            return operations;
        } catch (IOException e) {
            throw new DbSetupRuntimeException("failed to open " + builder.location, e);
        }
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
     * A builder to create a {@link Import} instance.
     *
     * @author sciencesakura
     */
    public static final class Builder {

        private final Map<String, Map<String, ValueGenerator<?>>> valueGenerators = new HashMap<>();

        private final URL location;

        private boolean built;

        private int left;

        private int top;

        private Builder(String location) {
            requireNonNull(location, "location must not be null");
            this.location = getClass().getClassLoader().getResource(location);
            if (this.location == null)
                throw new IllegalArgumentException(location + " not found");
        }

        /**
         * Constructs and returns a new {@link Import} instance.
         *
         * @return a new {@link Import} instance
         */
        @NotNull
        public Import build() {
            requireNotBuilt();
            built = true;
            return new Import(this);
        }

        /**
         * Specifies a start column index of worksheet to read data.
         * <p>
         * By default 0 is used.
         * </p>
         *
         * @param left a 0-based column index, must be non-negative
         * @return the reference to this object
         */
        @NotNull
        public Builder left(int left) {
            requireNotBuilt();
            if (left < 0)
                throw new IllegalArgumentException("left must be greater than or equal to 0");
            this.left = left;
            return this;
        }

        /**
         * Specifies a start row index of worksheet to read data.
         * <p>
         * By default 0 is used.
         * </p>
         *
         * @param top a 0-based row index, must be non-negative
         * @return the reference to this object
         */
        @NotNull
        public Builder top(int top) {
            requireNotBuilt();
            if (top < 0)
                throw new IllegalArgumentException("top must be greater than or equal to 0");
            this.top = top;
            return this;
        }

        /**
         * Add a value generator for the specified table and column.
         *
         * @param table          the table name
         * @param column         the column name
         * @param valueGenerator the generator
         * @return the reference to this object
         * @throws IllegalStateException if this builder has built an {@code Import} already
         */
        @NotNull
        public Builder withGeneratedValue(@NotNull String table, @NotNull String column,
                                          @NotNull ValueGenerator<?> valueGenerator) {
            requireNotBuilt();
            requireNonNull(table, "table must not be null");
            requireNonNull(column, "column must not be null");
            requireNonNull(valueGenerator, "valueGenerator must not be null");
            valueGenerators.computeIfAbsent(table, k -> new LinkedHashMap<>()).put(column, valueGenerator);
            return this;
        }

        private void requireNotBuilt() {
            if (built) throw new IllegalStateException("this operation has been built already");
        }
    }
}
