# Complete Excel LOOKUP Course: Beginner to Advanced

Welcome to your Excel lookup course! We'll learn how to find names, marks, employee salaries, product prices, and other information automatically using formulas.

You don't need any previous Excel knowledge. Each lesson includes a sample table, formula, explanation, practice exercise, and answer.

## Course roadmap

### 01

Excel basics for lookups

Cells, ranges, and references

### 02

LOOKUP

Basic and approximate searches

### 03

VLOOKUP

Vertical searches

### 04

HLOOKUP

Horizontal searches

### 05

XLOOKUP

Modern, flexible lookups

### 06

MATCH and XMATCH

Find the position of a value

### 07

INDEX + MATCH

Flexible lookups

### 08

Advanced techniques

Multiple criteria, wildcards, reverse searches, and two-way lookups

### 09

Error handling

IFERROR, IFNA, and missing results

### 10

Real-world projects

Student marks, employee data, and inventory

# Lesson 1 — Understand Excel basics

Before learning lookup functions, let's understand the structure of an Excel table.

Imagine that you have the following student records:

| A — ID | B — Name | C — Marks |
| ------ | -------- | --------- |
| 101    | Rahul    | 85        |
| 102    | Amit     | 92        |
| 103    | Priya    | 78        |
| 104    | Riya     | 95        |
| 105    | Arjun    | 88        |

- Cell: One box in Excel, such as `B2`.

- Column: A vertical group of cells, such as column B.

- Row: A horizontal group of cells, such as row 2.

- Range: A group of cells, such as `A2:C6`.

- Lookup value: The item you want to find, such as `Amit`.

- Return value: The information you want Excel to return, such as `92`.

For example, if you type `Amit` in cell `E2`, a formula can find his marks automatically.

# Lesson 2 — The LOOKUP function

The `LOOKUP` function finds a value in one range and returns a corresponding value from another range.

## Syntax

excel

```
=LOOKUP(lookup_value, lookup_vector, result_vector)
```

| Argument        | Meaning                     |
| --------------- | --------------------------- |
| `lookup_value`  | What you want to find       |
| `lookup_vector` | Where Excel searches        |
| `result_vector` | Where Excel gets the answer |

### Example: Find a student's grade

Suppose you have this table, with the minimum mark for each grade in ascending order.

| A — Minimum marks | B — Grade |
| ----------------- | --------- |
| 0                 | F         |
| 40                | D         |
| 50                | C         |
| 60                | B         |
| 70                | A         |
| 90                | A+        |

If a student's marks are in `D2`, enter:

excel

```
=LOOKUP(D2,A2:A7,B2:B7)
```

If `D2` contains `75`, the result is `A`.

How does it work? Excel finds the largest minimum-mark value that is less than or equal to 75, which is 70, and returns the corresponding grade.

Important: The lookup range for this kind of approximate `LOOKUP` must be sorted in ascending order. LOOKUP is generally not the best choice for finding an exact name or ID in an unsorted table.

## Lesson 3 — VLOOKUP

VLOOKUP means Vertical Lookup. It searches down the first column of a table and returns information from another column in the same row.

### Formula syntax

excel

```
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

The fourth argument is usually `FALSE` when you need an exact match.

### Example 1: Find Amit's marks

Use the student table from Lesson 1.

excel

```
=VLOOKUP("Amit",A2:C6,3,FALSE)
```

Formula result

# 92

Excel finds Amit in column B and returns the value from the third column of the selected range, column C.

### Understanding the column number

| Column in selected range | Column number |
| ------------------------ | ------------- |
| A — ID                   | 1             |
| B — Name                 | 2             |
| C — Marks                | 3             |

The column number starts at the left edge of the selected range, not necessarily column A of the worksheet.

### Example 2: Search using a cell reference

Suppose cell `E2` contains `Amit`.

excel

```
=VLOOKUP(E2,A2:C6,3,FALSE)
```

This formula returns `92`. If you change E2 to `Priya`, the result changes to `78`.

### Exact match versus approximate match

- `FALSE` — find an exact match. Use this for names, IDs, product codes, and employee numbers.

- `TRUE` — find an approximate match. The first column must be sorted in ascending order for reliable results.

### VLOOKUP limitations

- It searches only the first column of the selected table.

- It returns values from columns to the right.

- Inserting or deleting columns can make a hard-coded column number incorrect.

- It normally returns the first matching result.

### Practice 1

Using the student table in `A2:C6`, write a formula to find Riya's marks.

`=VLOOKUP("Riya",A2:C6,3,FALSE)`

`=VLOOKUP("Riya",A2:C6,2,FALSE)`

`=VLOOKUP("Riya",B2:C6,3,FALSE)`

Check answer

## Lesson 4 — HLOOKUP

HLOOKUP means Horizontal Lookup. It is useful when your data runs across columns instead of down rows.

### Sample table

|               | B     | C       | D      |
| ------------- | ----- | ------- | ------ |
| Row 1 — Name  | Rahul | Amit    | Priya  |
| Row 2 — Marks | 85    | 92      | 78     |
| Row 3 — City  | Delhi | Kolkata | Mumbai |

### Find Amit's marks

excel

```
=HLOOKUP("Amit",B1:D3,2,FALSE)
```

Result: `92`

How it works:

1. Search for Amit across the first row of the selected range.

2. Move down to the second row of that range.

3. Return the value at the matching column.

### Find Priya's city

excel

```
=HLOOKUP("Priya",B1:D3,3,FALSE)
```

Result: `Mumbai`.

Remember: VLOOKUP searches vertically in the first column; HLOOKUP searches horizontally in the first row.

### Practice 2

Using the same table, find Rahul's city. What is the formula?

Answer: `=HLOOKUP("Rahul",B1:D3,3,FALSE)` → Delhi

## Lesson 5 — XLOOKUP

XLOOKUP is one of the most useful modern Excel functions. It is easier to read and more flexible than VLOOKUP or HLOOKUP.

It is available in Microsoft 365 and newer Excel versions, but not in Excel 2016 or Excel 2019.

### Syntax

excel

```
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

| Argument       | Meaning                                       |
| -------------- | --------------------------------------------- |
| `lookup_value` | Value to find                                 |
| `lookup_array` | Range to search                               |
| `return_array` | Range containing the answer                   |
| `if_not_found` | Optional message if there is no match         |
| `match_mode`   | Exact, next smaller/larger, or wildcard match |
| `search_mode`  | Direction or search method                    |

### Example 1: Find Amit's marks

excel

```
=XLOOKUP("Amit",B2:B6,C2:C6)
```

Result: `92`.

Notice the difference: XLOOKUP uses a lookup range and a return range separately. You don't need to count the column number.

### Example 2: Display a message if the name is missing

excel

```
=XLOOKUP("Suman",B2:B6,C2:C6,"Not Found")
```

Result: `Not Found`.

### Example 3: Search from right to left

Suppose you want to find a name from the marks column, even though the name column is to its left.

excel

```
=XLOOKUP(92,C2:C6,B2:B6,"Not Found")
```

Result: `Amit`.

This works because XLOOKUP doesn't require the lookup column to be on the left.

### Example 4: Return the next smaller match

Using the sorted grade table from Lesson 2:

excel

```
=XLOOKUP(75,A2:A7,B2:B7,"Not Found",-1)
```

Result: `A`.

Here, match mode `-1` means exact match or next smaller item. With match mode `1`, Excel instead looks for an exact match or the next larger item.

### Example 5: Return the last matching record

If an employee appears multiple times in a transaction table, you may want the most recent matching row in the table's current order.

excel

```
=XLOOKUP("Amit",B2:B10,C2:C10,"Not Found",0,-1)
```

The final `-1` searches from the bottom upward, returning the last matching entry.

### XLOOKUP search modes

| Value | Meaning                                 |
| ----- | --------------------------------------- |
| `1`   | Search first to last (default)          |
| `-1`  | Search last to first                    |
| `2`   | Binary search on ascending-sorted data  |
| `-2`  | Binary search on descending-sorted data |

Use binary search only when the data is sorted correctly.

### Practice 3

If `E2` contains a student's name, write a formula to return their marks from `B2:B6` and `C2:C6`. Show `Name not found` if no match exists.

Answer: `=XLOOKUP(E2,B2:B6,C2:C6,"Name not found")`

## Lesson 6 — MATCH and XMATCH

These functions find the position of a value, not the value from a separate column.

### 6.1 MATCH

Syntax:

excel

```
=MATCH(lookup_value, lookup_array, [match_type])
```

Using the student names in `B2:B6`:

excel

```
=MATCH("Priya",B2:B6,0)
```

Result: `3`

Why? Priya is the third item in that range. MATCH returns the position within the range, not the worksheet row number.

The third argument means:

- `0` — exact match.

- `1` — approximate match, with the lookup range sorted ascending.

- `-1` — approximate match, with the lookup range sorted descending.

For most beginner tasks, use `0`.

### 6.2 XMATCH

XMATCH is a newer, more flexible alternative.

excel

```
=XMATCH("Priya",B2:B6)
```

Result: `3`.

XMATCH defaults to exact matching.

To find the last occurrence of Amit:

excel

```
=XMATCH("Amit",B2:B10,0,-1)
```

The final `-1` means search from last to first.

### 6.3 Wildcard matching

You can use wildcards when you know only part of a name.

excel

```
=XMATCH("Pri*",B2:B6,2)
```

This finds the first name starting with `Pri`, such as Priya. The `*` represents any number of characters; `?` represents one character.

| Wildcard | Meaning                  | Example                              |
| -------- | ------------------------ | ------------------------------------ |
| `*`      | Any number of characters | `Am*` matches Amit                   |
| `?`      | Exactly one character    | `Ri?a` matches Riya                  |
| `~`      | Escape a wildcard        | `~*` searches for a literal asterisk |

### Practice 4

What does this formula return?

excel

```
=MATCH("Amit",B2:B6,0)
```

Answer: `2`, because Amit is the second item in the selected range.

## Lesson 7 — INDEX + MATCH

This combination is especially useful when XLOOKUP is unavailable or when you want to separate the search and retrieval steps.

### 7.1 INDEX alone

Syntax:

excel

```
=INDEX(array,row_num,[column_num])
```

Example:

excel

```
=INDEX(C2:C6,3)
```

Result: `78`, the third value in the marks range.

### 7.2 Combine INDEX and MATCH

excel

```
=INDEX(C2:C6,MATCH("Amit",B2:B6,0))
```

Result: `92`.

Step by step:

1. `MATCH("Amit",B2:B6,0)` returns `2`.

2. `INDEX(C2:C6,2)` returns the second mark.

3. The second mark is `92`.

### 7.3 Search in either direction

Suppose your table has IDs in column A and names in column B. You can find an ID using a name:

excel

```
=INDEX(A2:A6,MATCH("Priya",B2:B6,0))
```

Result: `103`.

Unlike VLOOKUP, INDEX + MATCH can return a value from a column to the left of the lookup column.

### 7.4 Two-way lookup

Imagine a marks table with students in rows and subjects in columns.

| Student | Maths | Science | English |
| ------- | ----- | ------- | ------- |
| Rahul   | 85    | 90      | 78      |
| Amit    | 92    | 88      | 84      |
| Priya   | 78    | 95      | 91      |

Assume the table headers are in `A1:D4`. To find Amit's Science marks:

excel

```
=INDEX(B2:D4,MATCH("Amit",A2:A4,0),MATCH("Science",B1:D1,0))
```

Result: `88`.

The first MATCH identifies Amit's row. The second MATCH identifies the Science column. INDEX returns the value at their intersection.

### Practice 5

Using the same table, write a formula to find Priya's English marks.

Answer: `=INDEX(B2:D4,MATCH("Priya",A2:A4,0),MATCH("English",B1:D1,0))` → `91`

## Lesson 8 — Advanced lookup techniques

### 8.1 Multiple criteria

Suppose you have employee records and the same name may appear in different departments.

| A — Name | B — Department | C — Salary |
| -------- | -------------- | ---------- |
| Amit     | Sales          | 30000      |
| Amit     | IT             | 45000      |
| Priya    | HR             | 40000      |
| Rahul    | IT             | 50000      |

You want Amit's salary in IT.

In Excel versions that support dynamic arrays:

excel

```
=XLOOKUP(1,(A2:A5="Amit")*(B2:B5="IT"),C2:C5,"Not Found")
```

Result: `45000`.

The two logical tests produce TRUE/FALSE arrays. Multiplication converts TRUE to 1 and FALSE to 0; the row where both conditions are true produces 1.

You can also use cell references for the criteria, such as `E2` for the name and `F2` for the department:

excel

```
=XLOOKUP(1,(A2:A5=E2)*(B2:B5=F2),C2:C5,"Not Found")
```

### 8.2 Find a partial name

If a cell contains part of a name, wildcard matching can help.

excel

```
=XLOOKUP("*mit*",B2:B6,C2:C6,"Not Found",2)
```

This searches for text containing `mit`, such as Amit. The `2` enables wildcard matching.

### 8.3 Return multiple columns

XLOOKUP can return more than one column in modern Excel:

excel

```
=XLOOKUP(102,A2:A6,B2:C6,"Not Found")
```

This returns both Amit's name and marks, spilling the results into two adjacent cells.

Make sure the cells to the right of the formula are empty so the results can spill.

### 8.4 Find the nearest smaller value

Suppose the minimum thresholds are sorted ascending:

| Minimum sales | Commission |
| ------------- | ---------- |
| 0             | 0%         |
| 10000         | 2%         |
| 25000         | 5%         |
| 50000         | 8%         |

If sales are `32000` in `E2`:

excel

```
=XLOOKUP(E2,A2:A5,B2:B5,"Not Found",-1)
```

Result: `5%`.

The closest threshold that is less than or equal to 32000 is 25000.

### 8.5 Lookup using a two-dimensional grid

If you have both a row criterion and a column criterion, the INDEX + MATCH formula from Lesson 7 works well. XLOOKUP can also be nested:

excel

```
=XLOOKUP("Amit",A2:A4,XLOOKUP("Science",B1:D1,B2:D4))
```

Result: `88`.

The inner XLOOKUP selects the Science column; the outer XLOOKUP selects Amit's row from that returned column.

## Lesson 9 — Error handling and troubleshooting

Lookup formulas often encounter missing values. These functions help you manage them.

### IFERROR

excel

```
=IFERROR(VLOOKUP(E2,A2:C6,3,FALSE),"Not Found")
```

This displays `Not Found` if the lookup formula produces an error.

### IFNA

excel

```
=IFNA(XLOOKUP(E2,B2:B6,C2:C6),"Not Found")
```

IFNA handles the `#N/A` error specifically. It is useful when you want to handle a missing lookup without hiding other errors.

### Common lookup errors

| Error or problem         | Likely reason                               | What to do                                            |
| ------------------------ | ------------------------------------------- | ----------------------------------------------------- |
| `#N/A`                   | No match found                              | Check spelling, spaces, and data types                |
| `#REF!`                  | Invalid reference or column index           | Check the selected range and VLOOKUP column number    |
| `#VALUE!`                | Invalid argument or value                   | Review the formula arguments                          |
| Wrong approximate result | Data isn't sorted correctly                 | Sort the lookup range or use exact matching           |
| `#SPILL!`                | Results cannot spill into neighboring cells | Clear obstructing cells                               |
| Wrong row returned       | Duplicate lookup values                     | Use a unique ID or choose a first/last-match strategy |

Best practice: Use exact matching for IDs, names, and product codes unless you deliberately want an approximate result.

## Lesson 10 — Real-world Excel projects

Now let's apply the functions to practical situations.

### Project A: Student result sheet

Suppose your student database is in `A2:D6`.

| ID  | Name  | Marks | Grade |
| --- | ----- | ----- | ----- |
| 101 | Rahul | 85    | A     |
| 102 | Amit  | 92    | A+    |
| 103 | Priya | 78    | A     |
| 104 | Riya  | 95    | A+    |
| 105 | Arjun | 38    | F     |

If you enter a student ID in `F2`, return the student's name with:

excel

```
=XLOOKUP(F2,A2:A6,B2:B6,"ID not found")
```

Return the marks:

excel

```
=XLOOKUP(F2,A2:A6,C2:C6,"ID not found")
```

Return the grade:

excel

```
=XLOOKUP(F2,A2:A6,D2:D6,"ID not found")
```

This is useful for report cards and student databases.

### Project B: Product price lookup

Suppose you maintain this inventory:

| Product ID | Product    | Price |
| ---------- | ---------- | ----- |
| P101       | Keyboard   | 800   |
| P102       | Mouse      | 400   |
| P103       | Monitor    | 7500  |
| P104       | Headphones | 1200  |

If the product ID is entered in `E2`, use:

excel

```
=XLOOKUP(E2,A2:A5,C2:C5,"Product not found")
```

If E2 contains `P103`, the result is `7500`.

### Project C: Employee salary lookup

Suppose an employee database has:

| Employee ID | Name  | Department | Salary |
| ----------- | ----- | ---------- | ------ |
| E101        | Rahul | Sales      | 30000  |
| E102        | Amit  | IT         | 45000  |
| E103        | Priya | HR         | 40000  |
| E104        | Riya  | IT         | 50000  |

Find the salary using the employee ID in `F2`:

excel

```
=XLOOKUP(F2,A2:A5,D2:D5,"Employee not found")
```

This approach works well when employee IDs are unique.

## Complete lookup formula reference

| Function          | Example                                | Main purpose                      |
| ----------------- | -------------------------------------- | --------------------------------- |
| `LOOKUP`          | `=LOOKUP(E2,A2:A7,B2:B7)`              | Approximate lookup in sorted data |
| `VLOOKUP`         | `=VLOOKUP(E2,A2:C6,3,FALSE)`           | Vertical search                   |
| `HLOOKUP`         | `=HLOOKUP(E2,A1:D3,2,FALSE)`           | Horizontal search                 |
| `XLOOKUP`         | `=XLOOKUP(E2,A2:A6,C2:C6,"Not Found")` | Flexible lookup                   |
| `MATCH`           | `=MATCH(E2,B2:B6,0)`                   | Find a position                   |
| `XMATCH`          | `=XMATCH(E2,B2:B6)`                    | Modern position search            |
| `INDEX`           | `=INDEX(C2:C6,2)`                      | Return a value by position        |
| `INDEX` + `MATCH` | `=INDEX(C2:C6,MATCH(E2,B2:B6,0))`      | Flexible lookup                   |
| `IFERROR`         | `=IFERROR(formula,"Not Found")`        | Handle formula errors             |
| `IFNA`            | `=IFNA(formula,"Not Found")`           | Handle missing-match errors       |

## Final test — Check your knowledge

1. Which function is usually best for a modern exact lookup?

HLOOKUP

XLOOKUP

LOOKUP

2. What does MATCH return?

The matched value

The position of the match

The worksheet column name

3. Which VLOOKUP argument requests an exact match?

TRUE

FALSE

0 in the column number

4. Which formula can look to the left of the lookup column?

VLOOKUP only

INDEX + MATCH

HLOOKUP only

5. Which function searches from the bottom upward?

XLOOKUP with search mode -1

MATCH with match type 1

VLOOKUP with FALSE

Check all answers

## What should you learn next?

I recommend practising these three skills until you can write the formulas without looking at the examples:

1. VLOOKUP for compatibility with older Excel files.

2. XLOOKUP for everyday work in modern Excel.

3. INDEX + MATCH for flexible and two-way lookups.

After that, practise combining lookups with `IF`, `IFERROR`, `SUMIFS`, and Excel Tables to build complete reports.
