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
import org.jspecify.annotations.NonNull;

/**
 * An Operation which imports the Microsoft Excel file into the database.
 *
 * <p>We recommend to import {@code excel} method statically so that your code looks clearer.</p>
 * <pre>{@code import static com.sciencesakura.dbsetup.spreadsheet.Import.excel;}</pre>
 *
 * <p>Then you can use {@code excel} method as follows:</p>
 * <pre>{@code
 * var operation = excel("test-data.xlsx").build();
 * var dbSetup = new DbSetup(destination, operation);
 * dbSetup.launch();
 * }</pre>
 *
 * @author sciencesakura
 */
public final class Import implements Operation {

  /**
   * Create a new {@code Import.Builder} instance.
   *
   * @param location the {@code /}-separated path from classpath root to the Excel file
   * @return the new {@code Import.Builder} instance
   * @throws IllegalArgumentException if the Excel file is not found
   */
  @NonNull
  public static Builder excel(@NonNull String location) {
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

  /**
   * {@inheritDoc}
   */
  @Override
  public void execute(Connection connection, BinderConfiguration configuration) throws SQLException {
    internalOperation.execute(connection, configuration);
  }

  /**
   * A builder to create the {@code Import} operation.
   * The builder instance is created by the static method {@link Import#excel(String)}.
   * <table class="striped">
   *   <caption>Settings</caption>
   *   <thead>
   *     <tr>
   *       <th>Property</th>
   *       <th>Default Value</th>
   *       <th>To Customize</th>
   *     </tr>
   *   </thead>
   *   <tbody>
   *     <tr>
   *       <th>Sheets to include</th>
   *       <td>all sheets</td>
   *       <td>{@link #include(String...)} or {@link #include(Pattern...)}</td>
   *     </tr>
   *     <tr>
   *       <th>Sheets to exclude</th>
   *       <td>none</td>
   *       <td>{@link #exclude(String...)} or {@link #exclude(Pattern...)}</td>
   *     </tr>
   *     <tr>
   *       <th>Table name resolver</th>
   *       <td>Worksheet name is used as table name</td>
   *       <td>{@link #resolver(Map)} or {@link #resolver(Function)}</td>
   *     </tr>
   *     <tr>
   *       <th>Left margin</th>
   *       <td>{@code 0} columns</td>
   *       <td>{@link #left(int)} or {@link #margin(int, int)}</td>
   *     </tr>
   *     <tr>
   *       <th>Top margin</th>
   *       <td>{@code 0} rows</td>
   *       <td>{@link #top(int)} or {@link #margin(int, int)}</td>
   *     </tr>
   *     <tr>
   *       <th>Skip rows after header</th>
   *       <td>{@code 0} rows</td>
   *       <td>{@link #skipAfterHeader(int)}</td>
   *     </tr>
   *   </tbody>
   * </table>
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
     * Build a new {@code Import} operation instance.
     *
     * @return the new {@code Import} instance
     * @throws IllegalStateException if this builder has already built operation
     */
    @NonNull
    public Import build() {
      if (built) {
        throw new IllegalStateException("already built");
      }
      built = true;
      return new Import(this);
    }

    /**
     * Specifies a list of patterns to include.
     * The patterns are used to match the worksheet names in the Excel file.
     * By default, all worksheets are included.
     *
     * @param patterns the regular expressions to match worksheet names
     * @return the reference to this object
     */
    public Builder include(@NonNull String... patterns) {
      requireNonNull(patterns, "patterns must not be null");
      this.include = new Pattern[patterns.length];
      var i = 0;
      for (var pattern : patterns) {
        this.include[i++] = Pattern.compile(requireNonNull(pattern, "patterns must not contain null"));
      }
      return this;
    }

    /**
     * Specifies a list of patterns to include.
     * The patterns are used to match the worksheet names in the Excel file.
     * By default, all worksheets are included.
     *
     * @param patterns the regular expressions to match worksheet names
     * @return the reference to this object
     */
    public Builder include(@NonNull Pattern... patterns) {
      requireNonNull(patterns, "patterns must not be null");
      this.include = new Pattern[patterns.length];
      var i = 0;
      for (var pattern : patterns) {
        this.include[i++] = requireNonNull(pattern, "patterns must not contain null");
      }
      return this;
    }

    /**
     * Specifies a list of patterns to exclude.
     * The patterns are used to match the worksheet names in the Excel file.
     * By default, no worksheets are excluded.
     *
     * @param patterns the regular expressions to match worksheet names
     * @return the reference to this object
     */
    public Builder exclude(@NonNull String... patterns) {
      requireNonNull(patterns, "patterns must not be null");
      this.exclude = new Pattern[patterns.length];
      var i = 0;
      for (var pattern : patterns) {
        this.exclude[i++] = Pattern.compile(requireNonNull(pattern, "patterns must not contain null"));
      }
      return this;
    }

    /**
     * Specifies a list of patterns to exclude.
     * The patterns are used to match the worksheet names in the Excel file.
     * By default, no worksheets are excluded.
     *
     * @param patterns the regular expressions to match worksheet names
     * @return the reference to this object
     */
    public Builder exclude(@NonNull Pattern... patterns) {
      requireNonNull(patterns, "patterns must not be null");
      this.exclude = new Pattern[patterns.length];
      var i = 0;
      for (var pattern : patterns) {
        this.exclude[i++] = requireNonNull(pattern, "patterns must not contain null");
      }
      return this;
    }

    /**
     * Specifies a resolver to map worksheet names to table names.
     * By default, the worksheet name is used as the table name.
     *
     * @param resolver a map from worksheet name to table name
     * @return the reference to this object
     */
    public Builder resolver(@NonNull Map<String, String> resolver) {
      requireNonNull(resolver, "resolver must not be null");
      return resolver(resolver::get);
    }

    /**
     * Specifies a resolver to map worksheet names to table names.
     * By default, the worksheet name is used as the table name.
     *
     * @param resolver a map from worksheet name to table name
     * @return the reference to this object
     */
    public Builder resolver(@NonNull Function<String, String> resolver) {
      this.resolver = requireNonNull(resolver, "resolver must not be null");
      return this;
    }

    /**
     * Sets the left margin in columns.
     * By default, the left margin is {@code 0} columns.
     *
     * @param left the left margin in columns, must be non-negative
     * @return the reference to this object
     * @throws IllegalArgumentException if the argument is less than {@code 0}
     */
    public Builder left(int left) {
      if (left < 0) {
        throw new IllegalArgumentException("left must be greater than or equal to 0");
      }
      this.left = left;
      return this;
    }

    /**
     * Sets the top margin in rows.
     * By default, the top margin is {@code 0} rows.
     *
     * @param top the top margin in rows, must be non-negative
     * @return the reference to this object
     * @throws IllegalArgumentException if the argument is less than {@code 0}
     */
    public Builder top(int top) {
      if (top < 0) {
        throw new IllegalArgumentException("top must be greater than or equal to 0");
      }
      this.top = top;
      return this;
    }

    /**
     * Sets the left and top margins.
     * By default, the left margin is {@code 0} columns and the top margin is {@code 0} rows.
     *
     * @param left the left margin in columns, must be non-negative
     * @param top  the top margin in rows, must be non-negative
     * @return the reference to this object
     * @throws IllegalArgumentException if the arguments contain less than {@code 0}
     */
    public Builder margin(int left, int top) {
      return left(left).top(top);
    }

    /**
     * Sets the number of rows to skip after the header row.
     * By default, no rows are skipped after the header row.
     *
     * @param n the number of rows to skip after the header row, must be non-negative
     * @return the reference to this object
     * @throws IllegalArgumentException if the argument is less than {@code 0}
     */
    public Builder skipAfterHeader(int n) {
      if (n < 0) {
        throw new IllegalArgumentException("skipAfterHeader must be greater than or equal to 0");
      }
      this.skipAfterHeader = n;
      return this;
    }

    /**
     * Specifies a default value for the given table and column.
     *
     * @param table  the table name
     * @param column the column name to set the default value
     * @param value  the default value (nullable)
     * @return the reference to this object
     */
    public Builder withDefaultValue(@NonNull String table, @NonNull String column, Object value) {
      requireNonNull(table, "table must not be null");
      requireNonNull(column, "column must not be null");
      defaultValues.computeIfAbsent(table, k -> new LinkedHashMap<>()).put(column, value);
      return this;
    }

    /**
     * Specifies a value generator for the given table and column.
     * The value generator is used to generate a value for the column when inserting rows.
     *
     * @param table          the table name
     * @param column         the column name to set the value generator
     * @param valueGenerator the value generator to use
     * @return the reference to this object
     */
    public Builder withGeneratedValue(@NonNull String table, @NonNull String column,
                                      @NonNull ValueGenerator<?> valueGenerator) {
      requireNonNull(table, "table must not be null");
      requireNonNull(column, "column must not be null");
      requireNonNull(valueGenerator, "valueGenerator must not be null");
      valueGenerators.computeIfAbsent(table, k -> new LinkedHashMap<>()).put(column, valueGenerator);
      return this;
    }
  }
}
