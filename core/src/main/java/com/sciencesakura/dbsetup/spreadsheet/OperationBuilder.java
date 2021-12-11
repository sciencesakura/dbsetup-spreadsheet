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
import com.ninja_squad.dbsetup.Operations;
import com.ninja_squad.dbsetup.generator.ValueGenerator;
import com.ninja_squad.dbsetup.operation.Insert;
import com.ninja_squad.dbsetup.operation.Operation;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

final class OperationBuilder {

    private OperationBuilder() {
    }

    static Operation build(Import.Builder builder) {
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
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    throw new DbSetupRuntimeException("header row not found: " + sheetName + '[' + rowIndex + ']');
                }
                int width = row.getLastCellNum() - builder.left;
                if (width <= 0) {
                    throw new DbSetupRuntimeException("header row not found: " + sheetName + '[' + rowIndex + ']');
                }
                Insert.Builder ib = Insert.into(sheetName)
                        .columns(columns(row, builder.left, width, evaluator));
                Map<String, ValueGenerator<?>> valueGenerators = builder.valueGenerators.get(sheetName);
                if (valueGenerators != null) {
                    valueGenerators.forEach(ib::withGeneratedValue);
                }
                Map<String, ?> defaultValues = builder.defaultValues.get(sheetName);
                if (defaultValues != null) {
                    defaultValues.forEach(ib::withDefaultValue);
                }
                while ((row = sheet.getRow(++rowIndex)) != null) {
                    ib.values(values(row, builder.left, width, evaluator));
                }
                operations.add(ib.build());
            }
            return Operations.sequenceOf(operations);
        } catch (IOException e) {
            throw new DbSetupRuntimeException("failed to open " + builder.location, e);
        }
    }

    private static String a1(Cell cell) {
        return new CellReference(cell).formatAsString();
    }

    private static String a1(Sheet sheet, int r, int c) {
        return new CellReference(sheet.getSheetName(), r, c, false, false).formatAsString();
    }

    private static String[] columns(Row row, int left, int width, FormulaEvaluator evaluator) {
        String[] columns = new String[width];
        for (int i = 0; i < width; i++) {
            int c = left + i;
            Cell cell = row.getCell(c);
            if (cell == null) {
                throw new DbSetupRuntimeException("header cell not found: " + a1(row.getSheet(), row.getRowNum(), c));
            }
            Object value = value(cell, evaluator);
            if (value == null || "".equals(value)) {
                throw new DbSetupRuntimeException("header cell must not be blank: " + a1(cell));
            } else if (!(value instanceof String)) {
                throw new DbSetupRuntimeException("header cell must be string type: " + a1(cell));
            }
            columns[i] = (String) value;
        }
        return columns;
    }

    private static Object value(Cell cell, FormulaEvaluator evaluator) {
        switch (cell.getCellType()) {
            case NUMERIC:
                return DateUtil.isCellDateFormatted(cell) ? cell.getDateCellValue() : cell.getNumericCellValue();
            case STRING:
                return cell.getStringCellValue();
            case FORMULA:
                return value(evaluator.evaluateInCell(cell), evaluator);
            case BLANK:
                return null;
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case ERROR:
                throw new DbSetupRuntimeException("error value contained: " + a1(cell));
            default:
                throw new DbSetupRuntimeException("unsupported type: " + a1(cell));
        }
    }

    private static Object[] values(Row row, int left, int width, FormulaEvaluator evaluator) {
        Object[] values = new Object[width];
        for (int i = 0; i < width; i++) {
            Cell cell = row.getCell(left + i);
            values[i] = cell == null ? null : value(cell, evaluator);
        }
        return values;
    }
}
