// SPDX-License-Identifier: MIT

package com.sciencesakura.dbsetup.spreadsheet;

import static java.util.Objects.requireNonNull;

import com.ninja_squad.dbsetup.bind.BinderConfiguration;
import com.ninja_squad.dbsetup.generator.ValueGenerator;
import com.ninja_squad.dbsetup.operation.Operation;
import java.net.URL;
import java.sql.Connection;
import java.sql.SQLException;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.function.Function;
import java.util.regex.Pattern;
import org.jetbrains.annotations.NotNull;

/**
 * An Operation which imports the Microsoft Excel file into the tables.
 *
 * <h2>Usage</h2>
 * <p>When there are tables below:</p>
 * <pre>
 * {@code create table country (
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
 * }
 * </pre>
 * <p>
 * Create An Excel file with one worksheet per table, and name those worksheets
 * the same as the tables. If there is a dependency between the tables, put the
 * worksheet of the dependent table early.
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
 * Put the prepared Excel file on the classpath, and write code like the below:
 * </p>
 * <pre>
 * {@code import static com.sciencesakura.dbsetup.spreadsheet.Import.excel;
 *
 * var operation = excel("testdata.xlsx").build();
 * var dbSetup = new DbSetup(destination, operation);
 * dbSetup.launch();
 * }
 * </pre>
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
  public static Builder excel(@NotNull String location) {
    var urlLocation = Import.class.getClassLoader()
        .getResource(requireNonNull(location, "location must not be null"));
    if (urlLocation == null) {
      throw new IllegalArgumentException(location + " not found");
    }
    return new Builder(urlLocation);
  }

  private final Operation internalOperation;

  private Import(Builder builder) {
    this.internalOperation = OperationBuilder.build(builder);
  }

  @Override
  public void execute(Connection connection, BinderConfiguration configuration) throws SQLException {
    internalOperation.execute(connection, configuration);
  }

  /**
   * A builder to create the {@code Import} instance.
   *
   * <h2>Usage</h2>
   * <p>The default settings are:</p>
   * <ul>
   *   <li>{@code include([])}</li>
   *   <li>{@code exclude([])}</li>
   *   <li>{@code resolver(i -> i)}</li>
   *   <li>{@code left(0)}</li>
   *   <li>{@code top(0)}</li>
   *   <li>{@code margin(0, 0)}</li>
   *   <li>{@code skipAfterHeader(0)}</li>
   * </ul>
   *
   * @author sciencesakura
   */
  public static final class Builder {

    final URL location;
    Pattern[] include;
    Pattern[] exclude;
    Function<String, String> resolver = Function.identity();
    int left;
    int top;
    int skipAfterHeader;
    final Map<String, Map<String, Object>> defaultValues = new HashMap<>();
    final Map<String, Map<String, ValueGenerator<?>>> valueGenerators = new HashMap<>();
    private boolean built;

    private Builder(URL location) {
      this.location = location;
    }

    /**
     * Build a new {@code Import} instance.
     *
     * @return the new {@code Import} instance
     * @throws IllegalStateException if this method was called more than once on the same instance
     */
    public Import build() {
      if (built) {
        throw new IllegalStateException("already built");
      }
      built = true;
      return new Import(this);
    }

    /**
     * Specifies the patterns of the worksheet name to be included from the importing.
     *
     * @param patterns the regular expression patterns of the worksheet name to be included
     * @return the reference to this object
     */
    public Builder include(@NotNull String... patterns) {
      requireNonNull(patterns, "patterns must not be null");
      this.include = new Pattern[patterns.length];
      int i = 0;
      for (String pattern : patterns) {
        this.include[i++] = Pattern.compile(requireNonNull(pattern, "patterns must not contain null"));
      }
      return this;
    }

    /**
     * Specifies the patterns of the worksheet name to be included from the importing.
     *
     * @param patterns the regular expression patterns of the worksheet name to be included
     * @return the reference to this object
     */
    public Builder include(@NotNull Pattern... patterns) {
      requireNonNull(patterns, "patterns must not be null");
      this.include = new Pattern[patterns.length];
      int i = 0;
      for (Pattern pattern : patterns) {
        this.include[i++] = requireNonNull(pattern, "patterns must not contain null");
      }
      return this;
    }

    /**
     * Specifies the patterns of the worksheet name to be excluded from the importing.
     *
     * @param patterns the regular expression patterns of the worksheet name to be excluded
     * @return the reference to this object
     */
    public Builder exclude(@NotNull String... patterns) {
      requireNonNull(patterns, "patterns must not be null");
      this.exclude = new Pattern[patterns.length];
      int i = 0;
      for (String pattern : patterns) {
        this.exclude[i++] = Pattern.compile(requireNonNull(pattern, "patterns must not contain null"));
      }
      return this;
    }

    /**
     * Specifies the patterns of the worksheet name to be excluded from the importing.
     *
     * @param patterns the regular expression patterns of the worksheet name to be excluded
     * @return the reference to this object
     */
    public Builder exclude(@NotNull Pattern... patterns) {
      requireNonNull(patterns, "patterns must not be null");
      this.exclude = new Pattern[patterns.length];
      int i = 0;
      for (Pattern pattern : patterns) {
        this.exclude[i++] = requireNonNull(pattern, "patterns must not contain null");
      }
      return this;
    }

    /**
     * Specifies the resolver to map the worksheet name to the table name.
     *
     * @param resolver the resolver to map the worksheet name to the table name
     * @return the reference to this object
     */
    public Builder resolver(@NotNull Map<String, String> resolver) {
      requireNonNull(resolver, "resolver must not be null");
      return resolver(resolver::get);
    }

    /**
     * Specifies the resolver to map the worksheet name to the table name.
     *
     * @param resolver the resolver to map the worksheet name to the table name
     * @return the reference to this object
     */
    public Builder resolver(@NotNull Function<String, String> resolver) {
      this.resolver = requireNonNull(resolver, "resolver must not be null");
      return this;
    }

    /**
     * Specifies the start column index of the worksheet to read data.
     * <p>
     * By default {@code 0} is used.
     * </p>
     *
     * @param left the 0-based column index, must be non-negative
     * @return the reference to this object
     * @throws IllegalArgumentException if the specified value is negative
     */
    public Builder left(int left) {
      if (left < 0) {
        throw new IllegalArgumentException("left must be greater than or equal to 0");
      }
      this.left = left;
      return this;
    }

    /**
     * Specifies the start row index of the worksheet to read data.
     * <p>
     * By default {@code 0} is used.
     * </p>
     *
     * @param top the 0-based row index, must be non-negative
     * @return the reference to this object
     * @throws IllegalArgumentException if the specified value is negative
     */
    public Builder top(int top) {
      if (top < 0) {
        throw new IllegalArgumentException("top must be greater than or equal to 0");
      }
      this.top = top;
      return this;
    }

    /**
     * Specifies the start column index and row index of the worksheet to read data.
     * <p>
     * By default {@code (0, 0)} is used.
     * </p>
     *
     * @param left the 0-based column index, must be non-negative
     * @param top  the 0-based row index, must be non-negative
     * @return the reference to this object
     * @throws IllegalArgumentException if the specified value is negative
     */
    public Builder margin(int left, int top) {
      return left(left).top(top);
    }

    /**
     * Specifies the number of rows to skip after the header row.
     * <p>
     * By default {@code 0} is used.
     * </p>
     *
     * @param n the number of rows to skip after the header row, must be non-negative
     * @return the reference to this object
     * @throws IllegalArgumentException if the specified value is negative
     */
    public Builder skipAfterHeader(int n) {
      if (n < 0) {
        throw new IllegalArgumentException("skipAfterHeader must be greater than or equal to 0");
      }
      this.skipAfterHeader = n;
      return this;
    }

    /**
     * Specifies the default value for the given pair of table and column.
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
     * Specifies the value generator for the given pair of table and column.
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
