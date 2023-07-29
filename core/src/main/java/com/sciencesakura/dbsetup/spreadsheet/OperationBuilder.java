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
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

final class OperationBuilder {

    private OperationBuilder() {
    }

    static Operation build(Import.Builder builder) {
        try (Workbook workbook = WorkbookFactory.create(builder.location.openStream())) {
            List<Operation> operations = new ArrayList<>(workbook.getNumberOfSheets());
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                if (workbook.isSheetHidden(i) || workbook.isSheetVeryHidden(i)) {
                    continue;
                }
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();
                if (isExcluded(builder.include, builder.exclude, sheetName)) {
                    continue;
                }
                int rowIndex = builder.top;
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    throw new DbSetupRuntimeException("header row not found: " + sheetName + '[' + rowIndex + ']');
                }
                int width = row.getLastCellNum() - builder.left;
                if (width <= 0) {
                    throw new DbSetupRuntimeException("header row not found: " + sheetName + '[' + rowIndex + ']');
                }
                String tableName = builder.resolver.apply(sheetName);
                if (tableName == null) {
                    throw new DbSetupRuntimeException("could not resolve table name: " + sheetName);
                }
                Insert.Builder ib = Insert.into(tableName)
                        .columns(columns(row, builder.left, width, evaluator));
                setDefaultValues(ib, builder.defaultValues, tableName);
                setValueGenerators(ib, builder.valueGenerators, tableName);
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

    private static boolean isExcluded(Pattern[] include, Pattern[] exclude, String sheetName) {
        boolean included = include == null || include.length == 0;
        if (!included) {
            for (Pattern in : include) {
                if (in.matcher(sheetName).matches()) {
                    included = true;
                    break;
                }
            }
            if (!included) {
                return true;
            }
        }
        if (exclude == null) {
            return false;
        }
        for (Pattern ex : exclude) {
            if (ex.matcher(sheetName).matches()) {
                return true;
            }
        }
        return false;
    }

    private static void setDefaultValues(Insert.Builder ib,
                                         Map<String, Map<String, Object>> defaultValues,
                                         String tableName) {
        Map<String, ?> dv = defaultValues.get(tableName);
        if (dv == null) {
            return;
        }
        dv.forEach(ib::withDefaultValue);
    }

    private static void setValueGenerators(Insert.Builder ib,
                                           Map<String, Map<String, ValueGenerator<?>>> valueGenerators,
                                           String tableName) {
        Map<String, ValueGenerator<?>> vg = valueGenerators.get(tableName);
        if (vg == null) {
            return;
        }
        vg.forEach(ib::withGeneratedValue);
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
                throw new DbSetupRuntimeException("header cell must not be blank: " + a1(row.getSheet(), row.getRowNum(), c));
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
