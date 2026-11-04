# **Excel Lookup**

## Complete Course: Beginner to Advanced — **Microsoft 365**

We'll cover **all the major Excel lookup techniques step by step**, from traditional `VLOOKUP()` to advanced combinations of `XLOOKUP()`, `INDEX()`, `MATCH()`, `XMATCH()`, and `FILTER()`.

We'll use this practice workbook:

**[Download Excel Lookup Practice Dataset](https://docs.google.com/spreadsheets/d/1uSm7sp9royncTOcgbCHHnK-260UwtHum/edit?usp=sharing&ouid=102053457056127151946&rtpof=true&sd=true)**

Each lesson will contain explanations, syntax, worked examples, exercises, and a short assessment.

## Lesson 1: Understanding Lookup Functions

A lookup function searches for information and returns a related value.

Imagine you have this employee database:

| EmployeeID | EmployeeName | Department | City      | Salary |
| ---------- | ------------ | ---------- | --------- | -----: |
| E1001      | Aarav Das    | IT         | Kolkata   | 52,000 |
| E1002      | Priya Roy    | HR         | Hooghly   | 38,000 |
| E1003      | Rahul Sen    | Sales      | Howrah    | 32,000 |
| E1004      | Sneha Paul   | Finance    | Kolkata   | 45,000 |
| E1005      | Amit Ghosh   | IT         | Chinsurah | 58,000 |

Suppose you want to find the salary of employee `E1005`.

The answer is ₹58,000.

You could search manually, but Excel lookup functions can retrieve the information automatically.

Every lookup operation has three important components:

1. **Lookup value:** What you want to find.
2. **Lookup range:** Where Excel should search.
3. **Return value:** The information you want Excel to retrieve.

For example:

| Component    | Value             |
| ------------ | ----------------- |
| Lookup value | E1005             |
| Lookup range | Employee database |
| Return value | Salary            |
| Result       | ₹58,000           |

### Major lookup functions

| Function    | Primary purpose                                 |
| ----------- | ----------------------------------------------- |
| `VLOOKUP()` | Vertical lookup                                 |
| `HLOOKUP()` | Horizontal lookup                               |
| `XLOOKUP()` | Flexible modern lookup                          |
| `LOOKUP()`  | Approximate lookup in sorted data               |
| `INDEX()`   | Returns a value from a specified position       |
| `MATCH()`   | Finds the position of a value                   |
| `XMATCH()`  | Modern position-matching function               |
| `FILTER()`  | Returns all matching records                    |
| `CHOOSE()`  | Selects a value based on an index               |
| `OFFSET()`  | Creates a reference relative to a starting cell |

We'll also combine these functions with `IFNA()`, `IFERROR()`, `IF()`, and other functions to solve advanced problems.

---

## Lesson 2: VLOOKUP() — Vertical Lookup

### 2.1 What is VLOOKUP?

`VLOOKUP()` searches for a value in the first column of a table and returns information from another column in the same row.

### Syntax

```excel
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

| Argument        | Meaning                                                     |
| --------------- | ----------------------------------------------------------- |
| `lookup_value`  | The value to search for                                     |
| `table_array`   | The complete table range                                    |
| `col_index_num` | The position of the column containing the answer            |
| `range_lookup`  | `FALSE` for exact matching; `TRUE` for approximate matching |

### 2.2 Find an employee's name

In your workbook, open the **Employees** worksheet.

The employee IDs are in column A, names are in column B, and the dataset occupies rows 2–25.

Enter:

```excel
=VLOOKUP("E1005",A2:G25,2,FALSE)
```

**Result:** `Amit Ghosh`

Explanation:

- `"E1005"` is the lookup value.
- `A2:G25` is the lookup table.
- `2` specifies the second column of the selected table.
- `FALSE` requires an exact match.

### 2.3 Find the salary

```excel
=VLOOKUP("E1005",A2:G25,6,FALSE)
```

Result: `58000`

The salary is in column F, which is the sixth column of the selected range.

### 2.4 Use a cell reference

Suppose cell `I2` contains `E1005`.

```excel
=VLOOKUP(I2,$A$2:$G$25,6,FALSE)
```

The dollar signs lock the table range so it doesn't shift when you copy the formula.

### 2.5 Exact vs. approximate matching

**Exact match:**

```excel
=VLOOKUP("E1005",A2:G25,2,FALSE)
```

Use this for IDs, product codes, invoice numbers, and other unique identifiers.

**Approximate match:**

Suppose a grading table contains minimum marks and grades:

| Minimum Marks | Grade |
| ------------: | ----- |
|             0 | F     |
|            40 | D     |
|            50 | C     |
|            60 | B     |
|            70 | A     |
|            80 | A+    |
|            90 | O     |

If the table is in `A2:B8`, use:

```excel
=VLOOKUP(76,A2:B8,2,TRUE)
```

Result: `A`

Excel finds the greatest minimum mark that does not exceed 76, which is 70.

**Important:** For approximate `VLOOKUP()`, the first column should be sorted in ascending order.

### 2.6 Handle missing values

```excel
=IFNA(VLOOKUP("E9999",A2:G25,2,FALSE),"Employee Not Found")
```

Result: `Employee Not Found`

`IFNA()` handles the `#N/A` error when the lookup value isn't found.

### VLOOKUP limitation

`VLOOKUP()` normally searches in the first column of the selected range and returns a value from a column to its right.

It cannot directly search a name in column B and return the corresponding employee ID from column A.

We'll solve that problem with `XLOOKUP()` and `INDEX()` plus `MATCH()`.

---

## Lesson 3: HLOOKUP() — Horizontal Lookup

### 3.1 What is HLOOKUP?

`HLOOKUP()` searches across the first row of a table and returns a value from a specified row.

It's useful when data is arranged horizontally.

Open the **MonthlyPrices** worksheet in your workbook.

The data is arranged like this:

| Product  |    Jan |    Feb |    Mar |    Apr |    May |    Jun |
| -------- | -----: | -----: | -----: | -----: | -----: | -----: |
| Laptop   | 52,000 | 48,000 | 51,000 | 55,000 | 53,000 | 57,000 |
| Mouse    |    650 |    700 |    675 |    720 |    690 |    710 |
| Keyboard |  1,250 |  1,190 |  1,300 |  1,275 |  1,320 |  1,290 |

### Syntax

```excel
=HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])
```

### 3.2 Find the price of a product for a month

To retrieve the price of Mouse in May:

```excel
=HLOOKUP("May",B1:G6,3,FALSE)
```

Result: `690`

Explanation:

- `"May"` is the lookup value.
- `B1:G6` contains the month headings and price rows.
- `3` selects the third row of the selected range.
- `FALSE` requires an exact match.

### VLOOKUP vs. HLOOKUP

| Feature          | VLOOKUP                 | HLOOKUP              |
| ---------------- | ----------------------- | -------------------- |
| Search direction | Vertical                | Horizontal           |
| Searches         | First column            | First row            |
| Returns          | From a specified column | From a specified row |
| Best use         | Employee database       | Monthly price matrix |

Modern Excel users will often prefer `XLOOKUP()` for new workbooks because it is more flexible.

---

## Lesson 4: LOOKUP() — Traditional Lookup

The `LOOKUP()` function is useful for finding approximate matches in sorted data.

It has two main forms.

### Vector form

```excel
=LOOKUP(lookup_value, lookup_vector, [result_vector])
```

Example:

```excel
=LOOKUP(76,A2:A8,B2:B8)
```

If `A2:A8` contains ascending minimum marks and `B2:B8` contains the corresponding grades, this returns `A` for 76.

### Important characteristics

- It generally returns the value corresponding to the greatest lookup value less than or equal to the lookup value.
- The lookup vector should be sorted in ascending order.
- It is not the best choice for most exact-match ID lookups.
- It does not provide the same explicit match-mode control as `XLOOKUP()`.

For new workbooks, consider `XLOOKUP()` for greater control and readability.

---

## Lesson 5: XLOOKUP() — The Modern Lookup Function

`XLOOKUP()` is one of the most useful functions in modern Excel.

It can search in either direction, return values from a separate range, handle missing values, and support exact, approximate, and wildcard matching.

### Syntax

```excel
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

| Argument       | Meaning                                           |
| -------------- | ------------------------------------------------- |
| `lookup_value` | The value to find                                 |
| `lookup_array` | Where Excel searches                              |
| `return_array` | Where Excel retrieves the result                  |
| `if_not_found` | Optional result if no match exists                |
| `match_mode`   | Controls exact, approximate, or wildcard matching |
| `search_mode`  | Controls the search direction                     |

### 5.1 Find an employee's name

In the Employees worksheet:

```excel
=XLOOKUP("E1005",A2:A25,B2:B25)
```

Result: `Amit Ghosh`

Notice that the lookup range and return range are separate.

This is different from `VLOOKUP()`, which requires a table range and a column index.

### 5.2 Find an employee's salary

```excel
=XLOOKUP("E1005",A2:A25,F2:F25)
```

Result: `58000`

### 5.3 Perform a left lookup

Suppose you want to find an employee ID using their name.

```excel
=XLOOKUP("Ananya Bose",B2:B25,A2:A25)
```

Result: `E1008`

This works because `XLOOKUP()` can search one range and return a value from a range to its left.

### 5.4 Handle missing values

```excel
=XLOOKUP("E9999",A2:A25,B2:B25,"Employee Not Found")
```

Result: `Employee Not Found`

You don't need to wrap this formula in `IFNA()` just to handle a missing match.

### 5.5 Find the last matching record

Our Employees dataset contains repeated employee names.

For example, `Aarav Das` appears twice, with different employee IDs.

To return the last matching employee ID:

```excel
=XLOOKUP("Aarav Das",B2:B25,A2:A25,"Not Found",0,-1)
```

The final argument, `-1`, searches from the last entry toward the first.

This returns `E1021`.

**Lesson:** The search mode can be just as important as the lookup value.

---

## Lesson 6: INDEX() — Return a Value by Position

`INDEX()` returns a value from a specified position in a range.

### Syntax

```excel
=INDEX(array, row_num, [column_num])
```

Suppose cells `A2:A6` contain:

| Cell | EmployeeID |
| ---- | ---------- |
| A2   | E1001      |
| A3   | E1002      |
| A4   | E1003      |
| A5   | E1004      |
| A6   | E1005      |

Use:

```excel
=INDEX(A2:A6,3)
```

Result: `E1003`

Why?

Because `E1003` is the third item in the selected range.

The position is relative to the selected range, not the worksheet row number.

### Two-dimensional INDEX

Suppose `A1:C4` contains:

| ID    | Name      | Department |
| ----- | --------- | ---------- |
| E1001 | Aarav Das | IT         |
| E1002 | Priya Roy | HR         |
| E1003 | Rahul Sen | Sales      |

You can use:

```excel
=INDEX(A2:C4,2,3)
```

Result: `Sales`

The formula returns the value from the second row and third column of the selected range.

`INDEX()` becomes especially powerful when combined with `MATCH()`.

---

## Lesson 7: MATCH() — Find a Position

`MATCH()` returns the relative position of a value in a range.

### Syntax

```excel
=MATCH(lookup_value, lookup_array, [match_type])
```

For an exact match, use `0`.

Example:

```excel
=MATCH("E1005",A2:A25,0)
```

Result: `5`

Why?

Within `A2:A25`, `E1005` is the fifth item.

It is not returning worksheet row 6. It returns position 5 within the selected range.

### Find a student's position

In the StudentMarks worksheet:

```excel
=MATCH("ST108",A2:A13,0)
```

Result: `8`

`ST108` is the eighth item in the selected range.

### Match types

| Match Type | Meaning                                                             |
| ---------- | ------------------------------------------------------------------- |
| `0`        | Exact match                                                         |
| `1`        | Exact match or next smallest; lookup array must be sorted ascending |
| `-1`       | Exact match or next largest; lookup array must be sorted descending |

For normal ID searches, use `0`.

---

## Lesson 8: INDEX() + MATCH() — Flexible Lookup

This combination is an important intermediate-to-advanced technique.

`MATCH()` locates the position of the lookup value. `INDEX()` returns a value from that position.

### Example: Find an employee's salary

In the Employees worksheet:

```excel
=INDEX(F2:F25,MATCH("E1005",A2:A25,0))
```

Result: `58000`

### How it works

First, Excel evaluates:

```excel
=MATCH("E1005",A2:A25,0)
```

Result:

```text
5
```

Then Excel evaluates:

```excel
=INDEX(F2:F25,5)
```

Result:

```text
58000
```

The combined formula performs the lookup in one cell.

### Find an employee ID using their name

```excel
=INDEX(A2:A25,MATCH("Ananya Bose",B2:B25,0))
```

Result: `E1008`

Unlike traditional `VLOOKUP()`, this method can return a value from a column to the left of the search column.

### Why learn INDEX-MATCH if XLOOKUP exists?

`XLOOKUP()` is usually simpler in modern Excel. However, `INDEX()` and `MATCH()` remain valuable because:

- They're supported by many older Excel versions.
- They offer flexible lookup combinations.
- They help you understand how position-based lookups work.
- They're widely used in existing spreadsheets.

---

## Lesson 9: XMATCH() — Modern Position Matching

`XMATCH()` is the modern alternative to `MATCH()`.

### Syntax

```excel
=XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])
```

Example:

```excel
=XMATCH("P2010",Products!A2:A16,0)
```

Result: `10`

The product code `P2010` is the tenth item in the selected product-code range.

### Reverse search

```excel
=XMATCH("Aarav Das",Employees!B2:B25,0,-1)
```

Result: `21`

This searches from the bottom toward the top and returns the position of the last matching name.

### Match modes

| Mode | Meaning                          |
| ---- | -------------------------------- |
| `0`  | Exact match                      |
| `-1` | Exact match or next smaller item |
| `1`  | Exact match or next larger item  |
| `2`  | Wildcard match                   |

### Search modes

| Mode | Meaning                                 |
| ---- | --------------------------------------- |
| `1`  | First to last                           |
| `-1` | Last to first                           |
| `2`  | Binary search on ascending sorted data  |
| `-2` | Binary search on descending sorted data |

Binary search requires correctly sorted data; otherwise, results may be incorrect.

---

## Lesson 10: Two-Way Lookup

A two-way lookup finds a value using both a row heading and a column heading.

For example:

> Find the price of Monitor in April.

Open the MonthlyPrices worksheet.

We need to locate:

- The row containing `Monitor`.
- The column containing `Apr`.

### Using INDEX and MATCH

Assuming product names are in `A2:A6`, month headings are in `B1:G1`, and prices are in `B2:G6`:

```excel
=INDEX(B2:G6,MATCH("Monitor",A2:A6,0),MATCH("Apr",B1:G1,0))
```

Result: `15100`

### How it works

The first `MATCH()` finds the row containing Monitor.

The second `MATCH()` finds the column containing April.

`INDEX()` returns the value at their intersection.

### Using XLOOKUP inside XLOOKUP

```excel
=XLOOKUP("Monitor",A2:A6,XLOOKUP("Apr",B1:G1,B2:G6))
```

Result: `15100`

The inner `XLOOKUP()` selects the April price column. The outer `XLOOKUP()` retrieves the price for Monitor.

This is a powerful pattern for horizontal and vertical lookup combinations.

---

# Part 3 — Upper Intermediate Level

## Lesson 11: Lookup with Multiple Criteria

Sometimes one condition isn't enough.

Suppose you have these transaction records:

| EmployeeName | Department | City      | Salary |
| ------------ | ---------- | --------- | -----: |
| Priya Roy    | HR         | Hooghly   | 38,000 |
| Priya Roy    | IT         | Serampore | 60,000 |
| Rahul Sen    | Sales      | Howrah    | 32,000 |
| Rahul Sen    | Operations | Chinsurah | 38,500 |

You want to find Priya Roy's salary in the IT department.

Searching by name alone isn't sufficient because the name appears more than once.

### Method 1: XLOOKUP with multiple conditions

In the Transactions worksheet:

```excel
=XLOOKUP(1,(B2:B9="Priya Roy")*(C2:C9="IT"),E2:E9,"Not Found")
```

Result: `60000`

### Why does this work?

The first condition checks which rows contain `Priya Roy`.

The second condition checks which rows belong to the IT department.

Multiplying the Boolean arrays produces `1` only where both conditions are TRUE.

`XLOOKUP()` finds that row and returns the salary.

### Method 2: INDEX and MATCH

```excel
=INDEX(E2:E9,MATCH(1,(B2:B9="Priya Roy")*(C2:C9="IT"),0))
```

Result: `60000`

In Microsoft 365, this formula works directly without the old Ctrl+Shift+Enter array-formula method.

---

## Lesson 12: Wildcard Lookup

Wildcards allow you to search using partial text.

Excel supports two common wildcard characters:

| Wildcard | Meaning                      | Example                              |
| -------- | ---------------------------- | ------------------------------------ |
| `*`      | Any number of characters     | `Head*`                              |
| `?`      | Exactly one character        | `P200?`                              |
| `~`      | Escapes a wildcard character | `~*` searches for a literal asterisk |

### Example: Find a product beginning with "Head"

In the Products worksheet:

```excel
=XLOOKUP("Head*",B2:B16,A2:A16,"Not Found",2)
```

Result: `P2014`

Here:

- `B2:B16` contains product names.
- `A2:A16` contains product codes.
- Match mode `2` enables wildcard matching.

### Example: Find a product containing "Chair"

```excel
=XLOOKUP("*Chair*",B2:B16,A2:A16,"Not Found",2)
```

Result: `P2012`

**Important:** Wildcard matching with `XLOOKUP()` requires match mode `2`.

---

# Part 4 — Advanced Level

## Lesson 13: FILTER() — Return All Matching Results

Traditional lookup functions usually return one matching result.

But what if you want every employee in the IT department?

`FILTER()` can return multiple rows.

### Syntax

```excel
=FILTER(array, include, [if_empty])
```

### Example

In the Employees worksheet:

```excel
=FILTER(A2:G25,C2:C25="IT","No Employees Found")
```

This returns all employee records where the department is IT.

The results spill into adjacent cells automatically.

**Important:** Keep the output area clear so the results can spill. Otherwise, Excel may return `#SPILL!`.

### Filter by department and city

Find all IT employees in Kolkata:

```excel
=FILTER(A2:G25,(C2:C25="IT")*(D2:D25="Kolkata"),"No Match")
```

The multiplication operator means both conditions must be satisfied.

### Filter using OR conditions

Find employees in IT or HR:

```excel
=FILTER(A2:G25,((C2:C25="IT")+(C2:C25="HR"))>0,"No Match")
```

The addition operator combines the conditions.

---

## Lesson 14: Approximate Lookup and Range-Based Grading

Approximate matching is useful for:

- Student grades.
- Commission percentages.
- Discount slabs.
- Tax brackets.
- Shipping charges.

Suppose the grade thresholds are stored in `GradeScale!A2:B8`.

To find the grade for 76 marks:

### Using XLOOKUP

```excel
=XLOOKUP(76,GradeScale!A2:A8,GradeScale!B2:B8,"Not Found",-1)
```

Result: `A`

Match mode `-1` means exact match or next smaller item.

The minimum-marks column must be sorted in ascending order for this approach.

### Using VLOOKUP

```excel
=VLOOKUP(76,GradeScale!A2:B8,2,TRUE)
```

Result: `A`

Both methods are appropriate for this type of grading system when the table is structured correctly.

---

## Lesson 15: Lookup Across Worksheets

A lookup does not have to search the worksheet where the formula is written.

Suppose you're working in the Sales worksheet and want to retrieve the product name from Products.

The Sales worksheet stores the product code in column D.

### Retrieve the product name

```excel
=XLOOKUP(D2,Products!A$2:A$16,Products!B$2:B$16,"Not Found")
```

This formula:

1. Takes the product code from `D2`.
2. Searches the product-code range on the Products worksheet.
3. Returns the matching product name.

### Retrieve the unit price

```excel
=XLOOKUP(D2,Products!A$2:A$16,Products!E$2:E$16,"Not Found")
```

You can use this formula to retrieve the unit price for each sale.

### Calculate sales revenue

Revenue equals quantity multiplied by unit price.

Assuming the quantity is in `E2`:

```excel
=E2*XLOOKUP(D2,Products!A$2:A$16,Products!E$2:E$16,0)
```

This formula returns the line revenue for that sale.

For a real accounting workbook, consider handling missing product codes explicitly rather than treating every missing lookup as a zero price.

---

## Lesson 16: Return Multiple Columns with XLOOKUP

You can return multiple adjacent columns from a single lookup.

Suppose `I2` contains the product code `P2007`.

Enter:

```excel
=XLOOKUP(I2,Products!A2:A16,Products!B2:G16,"Not Found")
```

The result spills across multiple columns, returning the product name, category, supplier, price, stock quantity, and discount rate.

This is especially useful for building dynamic reports.

---

## Lesson 17: Return Multiple Matching Rows

Suppose the name `Priya Roy` appears more than once in the Transactions worksheet.

To retrieve every matching transaction:

```excel
=FILTER(A2:E9,B2:B9="Priya Roy","Not Found")
```

This returns all matching rows rather than only the first one.

### Combine FILTER with SORT

To return IT employees sorted by salary from highest to lowest:

```excel
=SORT(FILTER(A2:G25,C2:C25="IT"),6,-1)
```

The filter selects IT employees, and the sort orders the results by the sixth column of the filtered output, which is Salary.

---

## Lesson 18: CHOOSE() — Select Values by Position

`CHOOSE()` selects a value from a list based on an index number.

### Syntax

```excel
=CHOOSE(index_num,value1,[value2],...)
```

Example:

```excel
=CHOOSE(2,"VLOOKUP","XLOOKUP","INDEX-MATCH")
```

Result: `XLOOKUP`

The index `2` selects the second item.

### Use CHOOSE for a small lookup table

```excel
=CHOOSE(2,100,200,300)
```

Result: `200`

`CHOOSE()` is useful for selecting among a fixed set of alternatives, but it is not usually the best replacement for a lookup function when searching a large database.

---

## Lesson 19: OFFSET() — Reference-Based Lookup

`OFFSET()` returns a reference shifted from a starting cell by a specified number of rows and columns.

### Syntax

```excel
=OFFSET(reference,rows,cols,[height],[width])
```

Suppose:

- `A2:A6` contains product codes.
- `B2:B6` contains product names.
- `C2:C6` contains product prices.

To return the value two columns to the right of `A2`:

```excel
=OFFSET(A2,0,2)
```

Result: the value in `C2`.

`OFFSET()` is volatile, meaning Excel may recalculate it frequently. For ordinary lookups, `XLOOKUP()` and `INDEX()` are often clearer and more efficient.

---

# Part 5 — Error Handling and Formula Design

## Lesson 20: IFNA() vs. IFERROR()

Lookup formulas may return errors if a value is missing or the formula is incorrectly constructed.

### IFNA()

Handles only the `#N/A` error.

```excel
=IFNA(XLOOKUP("E9999",A2:A25,B2:B25),"Not Found")
```

### IFERROR()

Handles a broader range of errors.

```excel
=IFERROR(INDEX(F2:F25,MATCH("E9999",A2:A25,0)),"Lookup Failed")
```

### Which should you use?

| Function    | Best Use                                                             |
| ----------- | -------------------------------------------------------------------- |
| `IFNA()`    | When you want to handle missing lookup results specifically          |
| `IFERROR()` | When you deliberately want to handle a broader set of formula errors |

Avoid hiding every error automatically when debugging. A formula error may indicate a genuine problem that should be corrected.

---

# Part 6 — Professional Lookup Techniques

## Lesson 21: Lookup Using Excel Tables

Excel Tables make formulas easier to maintain as datasets grow.

Select your employee dataset and press **Ctrl + T**. Ensure the option **My table has headers** is selected.

If the table is named `Employees`, you can use structured references.

Example:

```excel
=XLOOKUP("E1005",Employees[EmployeeID],Employees[Salary],"Not Found")
```

This is easier to read than manually maintaining a fixed range.

New rows added to the table are automatically included in its structured references.

**Important:** The workbook already contains Excel Tables named `Employees`, `Products`, `Sales`, `Students`, `GradeScale`, and `Transactions`.

---

## Lesson 22: Lookup with Dynamic Column Selection

Sometimes you want the user to choose which field to retrieve.

Suppose:

- `I2` contains an Employee ID.
- `J1` contains a column heading such as `Salary`, `City`, or `Department`.

You can use `INDEX()` and `MATCH()` to retrieve the selected field:

```excel
=INDEX(A2:G25,MATCH(I2,A2:A25,0),MATCH(J1,A1:G1,0))
```

The formula:

1. Finds the employee's row.
2. Finds the selected column heading.
3. Returns the value at their intersection.

This is useful for creating interactive lookup tools and reports.

---

## Lesson 23: Lookup with Multiple Conditions Using Helper Columns

A helper column can simplify complicated lookups.

Suppose you need to identify an employee by both name and department.

In a new column, combine the two values:

```excel
=B2&"|"&C2
```

This creates a key such as:

```text
Priya Roy|IT
```

Then search the helper column:

```excel
=XLOOKUP("Priya Roy|IT",H2:H9,E2:E9,"Not Found")
```

This method is particularly useful when the conditions are complex or when you need to inspect the combined keys.

Choose a separator that cannot appear ambiguously in the source data, or use a more robust multi-condition formula.

---

# Final Project: Build an Employee Lookup System

Now combine the techniques you've learned.

Create a new worksheet called **Employee Lookup System**.

Set up this input area:

| Field         | Input or Result |
| ------------- | --------------- |
| Employee ID   | E1005           |
| Employee Name | Formula         |
| Department    | Formula         |
| City          | Formula         |
| Job Title     | Formula         |
| Salary        | Formula         |
| Status        | Formula         |

Assume the employee ID is entered in `B2`.

### Employee Name

```excel
=XLOOKUP(B2,Employees[EmployeeID],Employees[EmployeeName],"Not Found")
```

### Department

```excel
=XLOOKUP(B2,Employees[EmployeeID],Employees[Department],"Not Found")
```

### City

```excel
=XLOOKUP(B2,Employees[EmployeeID],Employees[City],"Not Found")
```

### Job Title

```excel
=XLOOKUP(B2,Employees[EmployeeID],Employees[JobTitle],"Not Found")
```

### Salary

```excel
=XLOOKUP(B2,Employees[EmployeeID],Employees[Salary],"Not Found")
```

### Status

```excel
=XLOOKUP(B2,Employees[EmployeeID],Employees[Status],"Not Found")
```

Change the employee ID in `B2` and observe how the entire lookup system updates.

This project demonstrates how lookup functions can automate real-world reporting tasks.

---

# Final Revision: Which Lookup Function Should You Use?

| Requirement                                 | Recommended Function                              |
| ------------------------------------------- | ------------------------------------------------- |
| Basic vertical lookup in an older workbook  | `VLOOKUP()`                                       |
| Horizontal lookup in a traditional layout   | `HLOOKUP()`                                       |
| Modern lookup in either direction           | `XLOOKUP()`                                       |
| Find a value's position                     | `MATCH()` or `XMATCH()`                           |
| Return a value by position                  | `INDEX()`                                         |
| Flexible lookup compatible with older Excel | `INDEX()` + `MATCH()`                             |
| Return every matching row                   | `FILTER()`                                        |
| Approximate lookup using sorted thresholds  | `XLOOKUP()` or `VLOOKUP()`                        |
| Find the last matching value                | `XLOOKUP()` with reverse search                   |
| Two-way lookup                              | `INDEX()` + `MATCH()` + `MATCH()`                 |
| Lookup using multiple conditions            | `XLOOKUP()` with Boolean conditions               |
| Handle missing lookup results               | `IFNA()` or `XLOOKUP()`'s `if_not_found` argument |
| Choose from a fixed list of values          | `CHOOSE()`                                        |
| Build a shifted cell reference              | `OFFSET()`                                        |
