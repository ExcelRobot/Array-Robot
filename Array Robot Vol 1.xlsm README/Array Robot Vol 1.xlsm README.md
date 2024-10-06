# <h1 id="oa-robot-definitions">OA Robot Definitions</h1>

\*\*Array Robot Vol 1.xlsm\*\* contains definitions for:

[112 Robot Commands](#command-definitions)<BR>[4 Robot Parameters](#parameter-definitions)<BR>[1 Robot Connection](#connection-definitions)<BR>[29 Robot Texts](#text-definitions)<BR>[3 Robot Images](#image-definitions)<BR>

<BR>

## Available Robot Commands

[Lambda](#lambda) | [Paste](#paste) | [Vol1](#vol1) | [WrapWith](#wrapwith) | [Other](#other)

### Lambda

| Name | Description |
| --- | --- |
| [Count Unique Values In Array](#count-unique-values-in-array) | Wrap with CountEachUniqueValue Lambda function. |
| [Fill Blanks From Above](#fill-blanks-from-above) | Wrap with FillFromAbove Lambda function. |
| [Fill Blanks From Below](#fill-blanks-from-below) | Wrap with FillFromBelow Lambda function. |
| [Fill Blanks From Left](#fill-blanks-from-left) | Wrap with FillFromLeft Lambda function. |
| [Fill Blanks From Right](#fill-blanks-from-right) | Wrap with FillFromRight Lambda function. |
| [If Blank Then Blank](#if-blank-then-blank) | Wrap with IFBLANK Then "" function so referenced blank cells will return blanks instead of zeros. |
| [If Blank Then NA](#if-blank-then-na) | Wrap with IFBLANK Then NA function. |
| [If Blank Then One (1)](#if-blank-then-one-1) | Wrap with ReplaceBlanksWithOnes Lambda function. |
| [If Blank Then Zero (0)](#if-blank-then-zero-0) | Wrap with ReplaceBlanksWithZeros Lambda function. |
| [If Zero (0) Then Blank](#if-zero-0-then-blank) | Wrap with ReplaceZerosWithBlanks Lambda function. |
| [Is One (1)?](#is-one-1) | Wraps active formula with Lambda function that returns TRUE for 1's and FALSE otherwise. |
| [Is Zero (0)?](#is-zero-0) | Wraps active formula with Lambda function that returns TRUE for 0's and FALSE if blank or non\-zero. |
| [Paste Count Unique Values In Array](#paste-count-unique-values-in-array) | Paste Counts of each unique value in the copied range. |
| [Remove Blank Rows And Columns](#remove-blank-rows-and-columns) | Wrap existing formula with RemoveBlanks Lambda function to remove blanks rows and columns from array. |
| [Reverse Columns (Flip Array Horizontally)](#reverse-columns-flip-array-horizontally) | Wrap existing formula with ReverseColumns Lambda function. |
| [Reverse Rows (Flip Array Vertically)](#reverse-rows-flip-array-vertically) | Wrap existing formula with ReverseRows Lambda function. |
| [Running Product Of Array](#running-product-of-array) | Wrap with RunningProduct Lambda function. |
| [Running Total Of Array](#running-total-of-array) | Wrap with RunningTotal Lambda function. |
| [Running Totals By Column](#running-totals-by-column) | Wrap with RunningTotalsByColumn Lambda function. |
| [Running Totals By Row](#running-totals-by-row) | Wrap with RunningTotalsByRow Lambda function. |

### Paste

| Name | Description |
| --- | --- |
| [Paste Sequence Across\/Right](#paste-sequence-acrossright) | Pastes a SEQUENCE function in the active cell spilling across to right based on the copied integer value. |
| [Paste Sequence Down](#paste-sequence-down) | Pastes a SEQUENCE function in the active cell spilling down based on the copied integer value. |

### Vol1

| Name | Description |
| --- | --- |
| [Are All True By Column?](#are-all-true-by-column) | Wrap with formula to return AND of each column. |
| [Are All True By Row?](#are-all-true-by-row) | Wrap with formula to return AND of each row. |
| [Are All True?](#are-all-true) | Wrap with formula with AND function. |
| [Are Some\/Any True By Column?](#are-someany-true-by-column) | Wrap with formula to return OR of each column. |
| [Are Some\/Any True By Row?](#are-someany-true-by-row) | Wrap with formula to return OR of each row. |
| [Are Some\/Any True?](#are-someany-true) | Wrap with formula with OR function. |
| [Average Of Array](#average-of-array) | Wrap with AVERAGE function. |
| [Average Of Array By Column](#average-of-array-by-column) | Wrap with formula to return average of each column. |
| [Average Of Array By Row](#average-of-array-by-row) | Wrap with formula to return average of each row. |
| [Concatenate Array](#concatenate-array) | Wrap with CONCAT function. |
| [Concatenate Array By Column](#concatenate-array-by-column) | Wrap with formula to return concatenation of each column. |
| [Concatenate Array By Row](#concatenate-array-by-row) | Wrap with formula to return concatenation of each row. |
| [Count Cells With Numbers](#count-cells-with-numbers) | Wrap with COUNT function to count the cells with numbers. |
| [Count Columns Of Array](#count-columns-of-array) | Wraps active array formula with COLUMNS function. |
| [Count Non\-Empty Cells](#count-non-empty-cells) | Wrap with COUNTA function to count the non\-empty cells. |
| [Count Rows Of Array](#count-rows-of-array) | Wraps active array formula with ROWS function. |
| [Count Unique Values In Array](#count-unique-values-in-array) | Wrap with CountEachUniqueValue Lambda function. |
| [Count Unique Values In Array By Column](#count-unique-values-in-array-by-column) | Wraps active array formula with functions to count unique values by column (not case sensitive). |
| [Count Unique Values In Array By Row](#count-unique-values-in-array-by-row) | Wraps active array formula with functions to count unique values by row (not case sensitive). |
| [Duplicate Array To Below](#duplicate-array-to-below) | Adds array formula below referencing the spill range of active array. |
| [Duplicate Array To Right](#duplicate-array-to-right) | Adds array formula to right referencing the spill range of active array. |
| [Extract Columns Of Array To Below](#extract-columns-of-array-to-below) | Adds array formula to below returning selected rows of active array. |
| [Extract Columns Of Array To Right](#extract-columns-of-array-to-right) | Adds array formula to right returning selected columns of active array. |
| [Extract Rows Of Array To Below](#extract-rows-of-array-to-below) | Adds array formula to below returning selected rows of active array. |
| [Extract Rows Of Array To Right](#extract-rows-of-array-to-right) | Adds array formula to right returning selected rows of active array. |
| [Fill Blanks From Above](#fill-blanks-from-above) | Wrap with FillFromAbove Lambda function. |
| [Fill Blanks From Below](#fill-blanks-from-below) | Wrap with FillFromBelow Lambda function. |
| [Fill Blanks From Left](#fill-blanks-from-left) | Wrap with FillFromLeft Lambda function. |
| [Fill Blanks From Right](#fill-blanks-from-right) | Wrap with FillFromRight Lambda function. |
| [If Blank Then Blank](#if-blank-then-blank) | Wrap with IFBLANK Then "" function so referenced blank cells will return blanks instead of zeros. |
| [If Blank Then NA](#if-blank-then-na) | Wrap with IFBLANK Then NA function. |
| [If Blank Then One (1)](#if-blank-then-one-1) | Wrap with ReplaceBlanksWithOnes Lambda function. |
| [If Blank Then Zero (0)](#if-blank-then-zero-0) | Wrap with ReplaceBlanksWithZeros Lambda function. |
| [If Error Then Blank](#if-error-then-blank) | Wrap with IFERROR function that returns "" if there is an error. |
| [If Error Then NA](#if-error-then-na) | Wrap with IFERROR Then NA function. |
| [If Error Then Zero (0)](#if-error-then-zero-0) | Wrap with IFERROR Then 0 function. |
| [If NA Then "NA"](#if-na-then-na) | Wrap with IFNA Then NA function. |
| [If NA Then Blank](#if-na-then-blank) | Wrap with IFNA Then "" function. |
| [If NA Then Zero (0)](#if-na-then-zero-0) | Wrap with IFNA Then 0 function. |
| [If Zero (0) Then Blank](#if-zero-0-then-blank) | Wrap with ReplaceZerosWithBlanks Lambda function. |
| [Is Error?](#is-error) | Wrap with ISERROR function. |
| [Is Even?](#is-even) | Wrap with ISEVEN function. |
| [Is False?](#is-false) | Wrap with NOT function. |
| [Is NA?](#is-na) | Wrap with ISNA function. |
| [Is Not True?](#is-not-true) | Wrap with NOT function. |
| [Is Odd?](#is-odd) | Wrap with ISODD function. |
| [Is One (1)?](#is-one-1) | Wraps active formula with Lambda function that returns TRUE for 1's and FALSE otherwise. |
| [Is Zero (0)?](#is-zero-0) | Wraps active formula with Lambda function that returns TRUE for 0's and FALSE if blank or non\-zero. |
| [Keep Columns Of Array](#keep-columns-of-array) | Wraps array formula with CHOOSECOLS of all selected columns. |
| [Keep First Column Of Array](#keep-first-column-of-array) | Wrap with formula to return first column of array. |
| [Keep First N Columns Of Array](#keep-first-n-columns-of-array) | Wrap with formula to return first N columns of array. |
| [Keep First N Rows Of Array](#keep-first-n-rows-of-array) | Wrap with formula to return first N rows of array. |
| [Keep First Row Of Array](#keep-first-row-of-array) | Wrap with formula to return first row of array. |
| [Keep Last Column Of Array](#keep-last-column-of-array) | Wrap with formula to return last column of array. |
| [Keep Last N Columns Of Array](#keep-last-n-columns-of-array) | Wrap with formula to return last N columns of array. |
| [Keep Last N Rows Of Array](#keep-last-n-rows-of-array) | Wrap with formula to return last N rows of array. |
| [Keep Last Row Of Array](#keep-last-row-of-array) | Wrap with formula to return last row of array. |
| [Keep Rows Of Array](#keep-rows-of-array) | Wraps array formula with CHOOSECOLS of all selected rows. |
| [Max Of Array](#max-of-array) | Wrap with MAX function to return the maximum. |
| [Max Of Array By Column](#max-of-array-by-column) | Wrap with formula to return max of each column. |
| [Max Of Array By Row](#max-of-array-by-row) | Wrap with formula to return max of each row. |
| [Median Of Array](#median-of-array) | Wrap with MEDIAN function to return the median. |
| [Min Of Array](#min-of-array) | Wrap with MIN function to return the minimum. |
| [Min Of Array By Column](#min-of-array-by-column) | Wrap with formula to return min of each column. |
| [Min Of Array By Row](#min-of-array-by-row) | Wrap with formula to return min of each row. |
| [Paste Count Unique Values In Array](#paste-count-unique-values-in-array) | Paste Counts of each unique value in the copied range. |
| [Paste Sequence Across\/Right](#paste-sequence-acrossright) | Pastes a SEQUENCE function in the active cell spilling across to right based on the copied integer value. |
| [Paste Sequence Down](#paste-sequence-down) | Pastes a SEQUENCE function in the active cell spilling down based on the copied integer value. |
| [Paste Unpivot Array](#paste-unpivot-array) | Pastes array formula with function to unpivot copied range by all column headers to the right. |
| [Product Of Array](#product-of-array) | Wrap with PRODUCT function. |
| [Product Of Array By Column](#product-of-array-by-column) | Wrap with formula to return product of each column. |
| [Product Of Array By Row](#product-of-array-by-row) | Wrap with formula to return product of each row. |
| [Remove Blank Rows And Columns](#remove-blank-rows-and-columns) | Wrap existing formula with RemoveBlanks Lambda function to remove blanks rows and columns from array. |
| [Remove First Column Of Array](#remove-first-column-of-array) | Wrap with DROP function to drop first column of array. |
| [Remove First N Columns Of Array](#remove-first-n-columns-of-array) | Wrap with DROP function to drop first N columns of array. |
| [Remove First N Rows Of Array](#remove-first-n-rows-of-array) | Wrap with DROP function to drop first N rows of array. |
| [Remove First Row Of Array](#remove-first-row-of-array) | Wrap with DROP function to drop first row of array. |
| [Remove Last Column Of Array](#remove-last-column-of-array) | Wrap with DROP function to drop last column of array. |
| [Remove Last N Columns Of Array](#remove-last-n-columns-of-array) | Wrap with DROP function to drop last N columns of array. |
| [Remove Last N Rows Of Array](#remove-last-n-rows-of-array) | Wrap with DROP function to drop last N rows of array. |
| [Remove Last Row Of Array](#remove-last-row-of-array) | Wrap with DROP function to remove last row of array. |
| [Remove Other Columns Of Array](#remove-other-columns-of-array) | Wraps array formula with CHOOSECOLS of all selected columns. |
| [Remove Other Rows Of Array](#remove-other-rows-of-array) | Wraps array formula with CHOOSECOLS of all selected rows. |
| [Reshape To One (1) Column](#reshape-to-one-1-column) | Wrap with TOCOL function to return as a column. |
| [Reshape To One (1) Row](#reshape-to-one-1-row) | Wrap with TOROW function to return as a row. |
| [Reverse Columns (Flip Array Horizontally)](#reverse-columns-flip-array-horizontally) | Wrap existing formula with ReverseColumns Lambda function. |
| [Reverse Rows (Flip Array Vertically)](#reverse-rows-flip-array-vertically) | Wrap existing formula with ReverseRows Lambda function. |
| [Running Product Of Array](#running-product-of-array) | Wrap with RunningProduct Lambda function. |
| [Running Total Of Array](#running-total-of-array) | Wrap with RunningTotal Lambda function. |
| [Running Totals By Column](#running-totals-by-column) | Wrap with RunningTotalsByColumn Lambda function. |
| [Running Totals By Row](#running-totals-by-row) | Wrap with RunningTotalsByRow Lambda function. |
| [Sort Array Asc](#sort-array-asc) | Wraps array formula with SORT function by selected columns of active array. |
| [Sort Array Desc](#sort-array-desc) | Wraps array formula with SORT function by selected columns of active array. |
| [Sum Array](#sum-array) | Wrap with SUM function. |
| [Sum Array By Column](#sum-array-by-column) | Wrap with formula to return sum of each column. |
| [Sum Array By Row](#sum-array-by-row) | Wrap with formula to return sum of each row. |
| [Transpose Array](#transpose-array) | Wrap with TRANSPOSE function. |
| [Unique Columns Of Array](#unique-columns-of-array) | Wrap with UNIQUE function to return the unique column values. |
| [Unique Rows Of Array](#unique-rows-of-array) | Wrap with UNIQUE function to return the unique values. |
| [Unpivot Array](#unpivot-array) | Wraps array formula with function to unpivot by all column headers to the right. |
| [Unpivot Array (By Selected Columns)](#unpivot-array-by-selected-columns) | Wraps array formula with function to unpivot by selected column headers. |

### WrapWith

| Name | Description |
| --- | --- |
| [Are All True By Column?](#are-all-true-by-column) | Wrap with formula to return AND of each column. |
| [Are All True By Row?](#are-all-true-by-row) | Wrap with formula to return AND of each row. |
| [Are All True?](#are-all-true) | Wrap with formula with AND function. |
| [Are Some\/Any True By Column?](#are-someany-true-by-column) | Wrap with formula to return OR of each column. |
| [Are Some\/Any True By Row?](#are-someany-true-by-row) | Wrap with formula to return OR of each row. |
| [Are Some\/Any True?](#are-someany-true) | Wrap with formula with OR function. |
| [Average Of Array](#average-of-array) | Wrap with AVERAGE function. |
| [Average Of Array By Column](#average-of-array-by-column) | Wrap with formula to return average of each column. |
| [Average Of Array By Row](#average-of-array-by-row) | Wrap with formula to return average of each row. |
| [Concatenate Array](#concatenate-array) | Wrap with CONCAT function. |
| [Concatenate Array By Column](#concatenate-array-by-column) | Wrap with formula to return concatenation of each column. |
| [Concatenate Array By Row](#concatenate-array-by-row) | Wrap with formula to return concatenation of each row. |
| [Count Cells With Numbers](#count-cells-with-numbers) | Wrap with COUNT function to count the cells with numbers. |
| [Count Non\-Empty Cells](#count-non-empty-cells) | Wrap with COUNTA function to count the non\-empty cells. |
| [Count Unique Values In Array](#count-unique-values-in-array) | Wrap with CountEachUniqueValue Lambda function. |
| [Fill Blanks From Above](#fill-blanks-from-above) | Wrap with FillFromAbove Lambda function. |
| [Fill Blanks From Below](#fill-blanks-from-below) | Wrap with FillFromBelow Lambda function. |
| [Fill Blanks From Left](#fill-blanks-from-left) | Wrap with FillFromLeft Lambda function. |
| [Fill Blanks From Right](#fill-blanks-from-right) | Wrap with FillFromRight Lambda function. |
| [If Blank Then Blank](#if-blank-then-blank) | Wrap with IFBLANK Then "" function so referenced blank cells will return blanks instead of zeros. |
| [If Blank Then NA](#if-blank-then-na) | Wrap with IFBLANK Then NA function. |
| [If Blank Then One (1)](#if-blank-then-one-1) | Wrap with ReplaceBlanksWithOnes Lambda function. |
| [If Blank Then Zero (0)](#if-blank-then-zero-0) | Wrap with ReplaceBlanksWithZeros Lambda function. |
| [If Error Then Blank](#if-error-then-blank) | Wrap with IFERROR function that returns "" if there is an error. |
| [If Error Then NA](#if-error-then-na) | Wrap with IFERROR Then NA function. |
| [If Error Then Zero (0)](#if-error-then-zero-0) | Wrap with IFERROR Then 0 function. |
| [If NA Then "NA"](#if-na-then-na) | Wrap with IFNA Then NA function. |
| [If NA Then Blank](#if-na-then-blank) | Wrap with IFNA Then "" function. |
| [If NA Then Zero (0)](#if-na-then-zero-0) | Wrap with IFNA Then 0 function. |
| [If Zero (0) Then Blank](#if-zero-0-then-blank) | Wrap with ReplaceZerosWithBlanks Lambda function. |
| [Is Error?](#is-error) | Wrap with ISERROR function. |
| [Is Even?](#is-even) | Wrap with ISEVEN function. |
| [Is False?](#is-false) | Wrap with NOT function. |
| [Is NA?](#is-na) | Wrap with ISNA function. |
| [Is Not True?](#is-not-true) | Wrap with NOT function. |
| [Is Odd?](#is-odd) | Wrap with ISODD function. |
| [Is One (1)?](#is-one-1) | Wraps active formula with Lambda function that returns TRUE for 1's and FALSE otherwise. |
| [Is Zero (0)?](#is-zero-0) | Wraps active formula with Lambda function that returns TRUE for 0's and FALSE if blank or non\-zero. |
| [Keep First Column Of Array](#keep-first-column-of-array) | Wrap with formula to return first column of array. |
| [Keep First N Columns Of Array](#keep-first-n-columns-of-array) | Wrap with formula to return first N columns of array. |
| [Keep First N Rows Of Array](#keep-first-n-rows-of-array) | Wrap with formula to return first N rows of array. |
| [Keep First Row Of Array](#keep-first-row-of-array) | Wrap with formula to return first row of array. |
| [Keep Last Column Of Array](#keep-last-column-of-array) | Wrap with formula to return last column of array. |
| [Keep Last N Columns Of Array](#keep-last-n-columns-of-array) | Wrap with formula to return last N columns of array. |
| [Keep Last N Rows Of Array](#keep-last-n-rows-of-array) | Wrap with formula to return last N rows of array. |
| [Keep Last Row Of Array](#keep-last-row-of-array) | Wrap with formula to return last row of array. |
| [Max Of Array](#max-of-array) | Wrap with MAX function to return the maximum. |
| [Max Of Array By Column](#max-of-array-by-column) | Wrap with formula to return max of each column. |
| [Max Of Array By Row](#max-of-array-by-row) | Wrap with formula to return max of each row. |
| [Median Of Array](#median-of-array) | Wrap with MEDIAN function to return the median. |
| [Min Of Array](#min-of-array) | Wrap with MIN function to return the minimum. |
| [Min Of Array By Column](#min-of-array-by-column) | Wrap with formula to return min of each column. |
| [Min Of Array By Row](#min-of-array-by-row) | Wrap with formula to return min of each row. |
| [Paste Count Unique Values In Array](#paste-count-unique-values-in-array) | Paste Counts of each unique value in the copied range. |
| [Product Of Array](#product-of-array) | Wrap with PRODUCT function. |
| [Product Of Array By Column](#product-of-array-by-column) | Wrap with formula to return product of each column. |
| [Product Of Array By Row](#product-of-array-by-row) | Wrap with formula to return product of each row. |
| [Remove Blank Rows And Columns](#remove-blank-rows-and-columns) | Wrap existing formula with RemoveBlanks Lambda function to remove blanks rows and columns from array. |
| [Reshape To One (1) Column](#reshape-to-one-1-column) | Wrap with TOCOL function to return as a column. |
| [Reshape To One (1) Row](#reshape-to-one-1-row) | Wrap with TOROW function to return as a row. |
| [Reverse Columns (Flip Array Horizontally)](#reverse-columns-flip-array-horizontally) | Wrap existing formula with ReverseColumns Lambda function. |
| [Reverse Rows (Flip Array Vertically)](#reverse-rows-flip-array-vertically) | Wrap existing formula with ReverseRows Lambda function. |
| [Running Product Of Array](#running-product-of-array) | Wrap with RunningProduct Lambda function. |
| [Running Total Of Array](#running-total-of-array) | Wrap with RunningTotal Lambda function. |
| [Running Totals By Column](#running-totals-by-column) | Wrap with RunningTotalsByColumn Lambda function. |
| [Running Totals By Row](#running-totals-by-row) | Wrap with RunningTotalsByRow Lambda function. |
| [Sum Array](#sum-array) | Wrap with SUM function. |
| [Sum Array By Column](#sum-array-by-column) | Wrap with formula to return sum of each column. |
| [Sum Array By Row](#sum-array-by-row) | Wrap with formula to return sum of each row. |
| [Transpose Array](#transpose-array) | Wrap with TRANSPOSE function. |
| [Unique Columns Of Array](#unique-columns-of-array) | Wrap with UNIQUE function to return the unique column values. |
| [Unique Rows Of Array](#unique-rows-of-array) | Wrap with UNIQUE function to return the unique values. |

### Other

| Name | Description |
| --- | --- |
| [Extract Numbers From Text To Columns](#extract-numbers-from-text-to-columns) | Given a text value, split any numbers from text into row array of numbers. |
| [If Text Then 0](#if-text-then-0) | Replaces text values with zeros. |
| [Import Array Robot Vol 1's Lambdas](#import-array-robot-vol-1s-lambdas) | Imports Array Robot Vol 1's lambda collection into active workbook. |
| [Paste Characters](#paste-characters) | Paste copied cells split by character using SplitByCharacter lambda. |
| [Paste Distinct\/Unique Row Values](#paste-distinctunique-row-values) | Pastes the values of distinct\/unique rows of the copied cells to the active cell. |
| [Paste Lookup](#paste-lookup) | Wrap active cell with XLOOKUP function to lookup each value in first row of copied cells and return corresponding last column of copied cells. |
| [Remove Columns Of Array](#remove-columns-of-array) | Removes specified columns of array using RemoveCols lambda. |
| [Remove Rows Of Array](#remove-rows-of-array) | Removes specified rows of array using RemoveRows lambda. |
| [Split By Character](#split-by-character) | Split active cell by character as row vector using SplitByCharacter lambda. |
| [Split By Digit](#split-by-digit) | Split active cell by digits as row vector using SplitByDigits lambda. |
| [Split Numbers From Text To Columns](#split-numbers-from-text-to-columns) | Given a text value, split any numbers from text into row array of numbers. |

<BR>

## Available Robot Parameters

| Name | Description |
| --- | --- |
| [Selected\_Column\_Indexes\_In\_Spilling\_Range](#selected_column_indexes_in_spilling_range) | List which columns of spilled array are included in selection. |
| [Selected\_Row\_Indexes\_In\_Spilling\_Range](#selected_row_indexes_in_spilling_range) | List which rows of spilled array are included in selection. |
| [Unselected\_Column\_Indexes\_In\_Spilling\_Range](#unselected_column_indexes_in_spilling_range) | List which columns of spilled array are not included in selection. |
| [Unselected\_Row\_Indexes\_In\_Spilling\_Range](#unselected_row_indexes_in_spilling_range) | List which rows of spilled array are not included in selection. |

<BR>

## Available Robot Connections

| Name | Description |
| --- | --- |
| [Lambda Robot](#lambda-robot) | Lambda Robot command collection. |

<BR>

## Available Robot Texts

| Name | Description |
| --- | --- |
| [Count Unique Values In Array.md](#count-unique-values-in-arraymd) | How to Use Count Unique Values In Array Lambda Command |
| [CountEachCharacter.lambda](#counteachcharacterlambda) | Counts how many of each character exists across all values in array. |
| [CountEachUniqueValue.lambda](#counteachuniquevaluelambda) | |
| [CROSSJOIN.lambda](#crossjoinlambda) | |
| [FillFromAbove.lambda](#fillfromabovelambda) | |
| [FillFromBelow.lambda](#fillfrombelowlambda) | |
| [FillFromLeft.lambda](#fillfromleftlambda) | |
| [FillFromRight.lambda](#fillfromrightlambda) | |
| [IFBLANK.lambda](#ifblanklambda) | |
| [IfText.lambda](#iftextlambda) | Definition of IfText lambda function. |
| [IsOne.lambda](#isonelambda) | |
| [IsZero.lambda](#iszerolambda) | |
| [RemoveBlanks.lambda](#removeblankslambda) | |
| [RemoveCols.lambda](#removecolslambda) | Definition of RemoveCols lambda function. |
| [RemoveRows.lambda](#removerowslambda) | Definition of RemoveRows lambda function. |
| [ReplaceBlanksWithOnes.lambda](#replaceblankswithoneslambda) | |
| [ReplaceBlanksWithZeros.lambda](#replaceblankswithzeroslambda) | |
| [ReplaceZerosWithBlanks.lambda](#replacezeroswithblankslambda) | |
| [ReverseColumns.lambda](#reversecolumnslambda) | |
| [ReverseRows.lambda](#reverserowslambda) | |
| [RunningProduct.lambda](#runningproductlambda) | |
| [RunningTotal.lambda](#runningtotallambda) | |
| [RunningTotalsByColumn.lambda](#runningtotalsbycolumnlambda) | |
| [RunningTotalsByRow.lambda](#runningtotalsbyrowlambda) | |
| [SplitByCharacter.lambda](#splitbycharacterlambda) | Lambda to split text into row vector of characters. |
| [SplitByDigit.lambda](#splitbydigitlambda) | Lambda to split text into row vector of characters. |
| [SplitNumbersFromText.lambda](#splitnumbersfromtextlambda) | Lambda to split numbers out of text as separate columns. |
| [TILE.lambda](#tilelambda) | |
| [UNPIVOT.lambda](#unpivotlambda) | |

<BR>

## Available Robot Images

| Name | Description |
| --- | --- |
| [Count\_Unique\_Example.png](#count_unique_examplepng) | |
| [MyImage.gif](#myimagegif) | |
| [Remove\_Columns\_Rows\_Of\_Array.png](#remove_columns_rows_of_arraypng) | Remove Columns\/Rows Of Array |

<BR>

## Command Definitions

<BR>

### Are All True By Column?

*Wrap with formula to return AND of each column.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=BYCOL(\[\[ActiveCell::Formula\]\],LAMBDA(x,AND(x)))</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>at</code></li><li><code>atc</code></li><li><code>atbc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Are All True By Row?

*Wrap with formula to return AND of each row.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=BYROW(\[\[ActiveCell::Formula\]\],LAMBDA(x,AND(x)))</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>at</code></li><li><code>atr</code></li><li><code>atbr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Are All True?

*Wrap with formula with AND function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=AND(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <code>and</code> |

[^Top](#oa-robot-definitions)

<BR>

### Are Some\/Any True By Column?

*Wrap with formula to return OR of each column.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=BYCOL(\[\[ActiveCell::Formula\]\],LAMBDA(x,OR(x)))</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>st</code></li><li><code>stc</code></li><li><code>stbc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Are Some\/Any True By Row?

*Wrap with formula to return OR of each row.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=BYROW(\[\[ActiveCell::Formula\]\],LAMBDA(x,OR(x)))</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>st</code></li><li><code>str</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Are Some\/Any True?

*Wrap with formula with OR function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=OR(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <code>or</code> |

[^Top](#oa-robot-definitions)

<BR>

### Average Of Array

*Wrap with AVERAGE function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=AVERAGE(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <ol><li><code>a</code></li><li><code>av</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Average Of Array By Column

*Wrap with formula to return average of each column.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=BYCOL(\[\[ActiveCell::Formula\]\],LAMBDA(x,AVERAGE(x)))</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>ac</code></li><li><code>abc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Average Of Array By Row

*Wrap with formula to return average of each row.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=BYROW(\[\[ActiveCell::Formula\]\],LAMBDA(x,AVERAGE(x)))</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>ar</code></li><li><code>abr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Concatenate Array

*Wrap with CONCAT function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=CONCAT(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <ol><li><code>c</code></li><li><code>ca</code></li><li><code>con</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Concatenate Array By Column

*Wrap with formula to return concatenation of each column.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=BYCOL(\[\[ActiveCell::Formula\]\],LAMBDA(x,CONCAT(x)))</code> |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>cc</code></li><li><code>cbc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Concatenate Array By Row

*Wrap with formula to return concatenation of each row.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=BYROW(\[\[ActiveCell::Formula\]\],LAMBDA(x,CONCAT(x)))</code> |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>cr</code></li><li><code>cbr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Count Cells With Numbers

*Wrap with COUNT function to count the cells with numbers.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=COUNT(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent |

[^Top](#oa-robot-definitions)

<BR>

### Count Columns Of Array

*Wraps active array formula with COLUMNS function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=COLUMNS(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <code>cc</code> |

[^Top](#oa-robot-definitions)

<BR>

### Count Non\-Empty Cells

*Wrap with COUNTA function to count the non\-empty cells.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=COUNTA(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent |

[^Top](#oa-robot-definitions)

<BR>

### Count Rows Of Array

*Wraps active array formula with ROWS function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=ROWS(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <code>cr</code> |

[^Top](#oa-robot-definitions)

<BR>

### Count Unique Values In Array

*Wrap with CountEachUniqueValue Lambda function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

[How to Use Count Unique Values In Array Lambda Command](<.\Documentation\Count Unique Values In Array.md>)

> \*\*Note:\*\* Great for using in competitions to answer questions like "which brand is mentioned the most in the data?"

| Property | Value |
| --- | --- |
| Formula | <code>\=CountEachUniqueValue(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | [CountEachUniqueValue.lambda](#counteachuniquevaluelambda) |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Documentation | [Count Unique Values In Array.md](#count-unique-values-in-arraymd) |
| Launch Codes | <code>cu</code> |

[^Top](#oa-robot-definitions)

<BR>

### Count Unique Values In Array By Column

*Wraps active array formula with functions to count unique values by column (not case sensitive).*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=BYCOL(\[\[ActiveCell::Formula\]\],LAMBDA(x,COUNTA(UNIQUE(TOCOL(x)))))</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>cuc</code></li><li><code>cubc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Count Unique Values In Array By Row

*Wraps active array formula with functions to count unique values by row (not case sensitive).*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=BYROW(\[\[ActiveCell::Formula\]\],LAMBDA(x,COUNTA(UNIQUE(TOCOL(x)))))</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>cur</code></li><li><code>cubr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Duplicate Array To Below

*Adds array formula below referencing the spill range of active array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=\[\[ActiveCell.SpillParent.SpillingToRange\]\]</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent.SpillingToRange.Rows(1).AdjacentBlank.Cells(1,1)\]\]</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>da</code></li><li><code>dab</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Duplicate Array To Right

*Adds array formula to right referencing the spill range of active array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=\[\[ActiveCell.SpillParent.SpillingToRange\]\]</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent.SpillingToRange.Columns(1).AdjacentBlank.Cells(1,1)\]\]</code> |
| Keyboard Shortcut | <code>Ctrl + Shift + r</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>da</code></li><li><code>dar</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Extract Columns Of Array To Below

*Adds array formula to below returning selected rows of active array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=CHOOSECOLS(\[\[ActiveCell.SpillParent.SpillingToRange\]\],{{[Selected\_Column\_Indexes\_In\_Spilling\_Range](#selected_column_indexes_in_spilling_range)}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent.SpillingToRange.Rows(1).AdjacentBlank.Cells(1,1)\]\]</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>ec</code></li><li><code>ecb</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Extract Columns Of Array To Right

*Adds array formula to right returning selected columns of active array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=CHOOSECOLS(\[\[ActiveCell.SpillParent.SpillingToRange\]\],{{[Selected\_Column\_Indexes\_In\_Spilling\_Range](#selected_column_indexes_in_spilling_range)}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent.SpillingToRange.Columns(1).AdjacentBlank.Cells(1,1)\]\]</code> |
| Keyboard Shortcut | <code>ctrl + alt + r</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>ec</code></li><li><code>ecr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Extract Numbers From Text To Columns

*Given a text value, split any numbers from text into row array of numbers.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SplitNumbersFromText(\[\[ActiveCell\]\])</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent.SpillingToRange.Columns(1).AdjacentBlank.Cells(1,1)\]\]</code> |
| Formula Dependencies | [SplitNumbersFromText.lambda](#splitnumbersfromtextlambda) |
| Launch Codes | <code>en</code> |

[^Top](#oa-robot-definitions)

<BR>

### Extract Rows Of Array To Below

*Adds array formula to below returning selected rows of active array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=CHOOSEROWS(\[\[ActiveCell.SpillParent.SpillingToRange\]\],{{[Selected\_Row\_Indexes\_In\_Spilling\_Range](#selected_row_indexes_in_spilling_range)}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent.SpillingToRange.Rows(1).AdjacentBlank.Cells(1,1)\]\]</code> |
| Keyboard Shortcut | <code>ctrl + alt + d</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>er</code></li><li><code>erb</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Extract Rows Of Array To Right

*Adds array formula to right returning selected rows of active array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=CHOOSEROWS(\[\[ActiveCell.SpillParent.SpillingToRange\]\],{{[Selected\_Row\_Indexes\_In\_Spilling\_Range](#selected_row_indexes_in_spilling_range)}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent.SpillingToRange.Columns(1).AdjacentBlank.Cells(1,1)\]\]</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>er</code></li><li><code>err</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Fill Blanks From Above

*Wrap with FillFromAbove Lambda function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=FillFromAbove(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | <ol><li>[TILE.lambda](#tilelambda)</li><li>[FillFromAbove.lambda](#fillfromabovelambda)</li></ol> |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>fb</code></li><li><code>fba</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Fill Blanks From Below

*Wrap with FillFromBelow Lambda function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=FillFromBelow(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | <ol><li>[TILE.lambda](#tilelambda)</li><li>[FillFromBelow.lambda](#fillfrombelowlambda)</li></ol> |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>fb</code></li><li><code>fbb</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Fill Blanks From Left

*Wrap with FillFromLeft Lambda function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=FillFromLeft(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | <ol><li>[TILE.lambda](#tilelambda)</li><li>[FillFromLeft.lambda](#fillfromleftlambda)</li></ol> |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>fb</code></li><li><code>fbl</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Fill Blanks From Right

*Wrap with FillFromRight Lambda function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=FillFromRight(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | <ol><li>[TILE.lambda](#tilelambda)</li><li>[FillFromRight.lambda](#fillfromrightlambda)</li></ol> |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>fb</code></li><li><code>fbr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### If Blank Then Blank

*Wrap with IFBLANK Then "" function so referenced blank cells will return blanks instead of zeros.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=IFBLANK(\[\[ActiveCell::Formula\]\],"")</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | [IFBLANK.lambda](#ifblanklambda) |
| User Context Filter | ExcelActiveCellContainsFormula |
| Launch Codes | <code>bb</code> |

[^Top](#oa-robot-definitions)

<BR>

### If Blank Then NA

*Wrap with IFBLANK Then NA function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=IFBLANK(\[\[ActiveCell::Formula\]\],"NA")</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | [IFBLANK.lambda](#ifblanklambda) |
| User Context Filter | ExcelActiveCellContainsFormula |
| Launch Codes | <code>bna</code> |

[^Top](#oa-robot-definitions)

<BR>

### If Blank Then One (1)

*Wrap with ReplaceBlanksWithOnes Lambda function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=ReplaceBlanksWithOnes(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | [ReplaceBlanksWithOnes.lambda](#replaceblankswithoneslambda) |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <code>b1</code> |

[^Top](#oa-robot-definitions)

<BR>

### If Blank Then Zero (0)

*Wrap with ReplaceBlanksWithZeros Lambda function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=ReplaceBlanksWithZeros(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | [ReplaceBlanksWithZeros.lambda](#replaceblankswithzeroslambda) |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <code>b0</code> |

[^Top](#oa-robot-definitions)

<BR>

### If Error Then Blank

*Wrap with IFERROR function that returns "" if there is an error.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=IFERROR(\[\[ActiveCell::Formula\]\],"")</code> |
| Destination Range Address | <code>\[\[ActiveCell\]\]</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellContainsFormula |
| Launch Codes | <code>eb</code> |

[^Top](#oa-robot-definitions)

<BR>

### If Error Then NA

*Wrap with IFERROR Then NA function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=IFERROR(\[\[ActiveCell::Formula\]\],"NA")</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellContainsFormula |
| Launch Codes | <ol><li><code>ena</code></li><li><code>ife</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### If Error Then Zero (0)

*Wrap with IFERROR Then 0 function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=IFERROR(\[\[ActiveCell::Formula\]\],0)</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellContainsFormula |
| Launch Codes | <ol><li><code>e0</code></li><li><code>ife</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### If NA Then "NA"

*Wrap with IFNA Then NA function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=IFNA(\[\[ActiveCell::Formula\]\],"NA")</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellContainsFormula |
| Launch Codes | <ol><li><code>nana</code></li><li><code>ifna</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### If NA Then Blank

*Wrap with IFNA Then "" function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=IFNA(\[\[ActiveCell::Formula\]\],"")</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellContainsFormula |
| Launch Codes | <code>nab</code> |

[^Top](#oa-robot-definitions)

<BR>

### If NA Then Zero (0)

*Wrap with IFNA Then 0 function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=IFNA(\[\[ActiveCell::Formula\]\],0)</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellContainsFormula |
| Launch Codes | <ol><li><code>na0</code></li><li><code>ifna</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### If Text Then 0

*Replaces text values with zeros.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=IfText(\[\[ActiveCell::Formula\]\], 0)</code> |
| Formula Dependencies | [IfText.lambda](#iftextlambda) |
| Launch Codes | <code>t0</code> |

[^Top](#oa-robot-definitions)

<BR>

### If Zero (0) Then Blank

*Wrap with ReplaceZerosWithBlanks Lambda function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=ReplaceZerosWithBlanks(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | [ReplaceZerosWithBlanks.lambda](#replacezeroswithblankslambda) |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellContainsFormula |
| Launch Codes | <ol><li><code>0b</code></li><li><code>if0</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Import Array Robot Vol 1's Lambdas

*Imports Array Robot Vol 1's lambda collection into active workbook.*

<sup>`@Array Robot Vol 1.xlsm` `!VBA Macro Command` </sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modImportLambdas.ImportAllLambdas](./VBA/modImportLambdas.bas)("Array Robot Vol 1.xlsm")</code> |
| Macro Workbook Connection | [Lambda Robot](#lambda-robot) |

[^Top](#oa-robot-definitions)

<BR>

### Is Error?

*Wrap with ISERROR function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=ISERROR(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellContainsFormula |
| Launch Codes | <ol><li><code>ie</code></li><li><code>ise</code></li><li><code>ier</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Is Even?

*Wrap with ISEVEN function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=ISEVEN(+\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellValueIsInteger |
| Launch Codes | <ol><li><code>ie</code></li><li><code>ise</code></li><li><code>iev</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Is False?

*Wrap with NOT function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=NOT(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellContainsFormula |
| Launch Codes | <ol><li><code>if</code></li><li><code>isf</code></li><li><code>ifa</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Is NA?

*Wrap with ISNA function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=ISNA(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellContainsFormula |
| Launch Codes | <ol><li><code>ina</code></li><li><code>isna</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Is Not True?

*Wrap with NOT function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=NOT(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellContainsFormula |
| Launch Codes | <code>not</code> |

[^Top](#oa-robot-definitions)

<BR>

### Is Odd?

*Wrap with ISODD function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=ISODD(+\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellValueIsInteger |
| Launch Codes | <code>odd</code> |

[^Top](#oa-robot-definitions)

<BR>

### Is One (1)?

*Wraps active formula with Lambda function that returns TRUE for 1's and FALSE otherwise.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=IsOne(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | [IsOne.lambda](#isonelambda) |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellContainsFormula |
| Launch Codes | <code>is1</code> |

[^Top](#oa-robot-definitions)

<BR>

### Is Zero (0)?

*Wraps active formula with Lambda function that returns TRUE for 0's and FALSE if blank or non\-zero.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=IsZero(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | [IsZero.lambda](#iszerolambda) |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellContainsFormula |
| Launch Codes | <code>is0</code> |

[^Top](#oa-robot-definitions)

<BR>

### Keep Columns Of Array

*Wraps array formula with CHOOSECOLS of all selected columns.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=CHOOSECOLS(\[\[ActiveCell.SpillParent::Formula\]\],{{[Selected\_Column\_Indexes\_In\_Spilling\_Range](#selected_column_indexes_in_spilling_range)}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <code>kc</code> |

[^Top](#oa-robot-definitions)

<BR>

### Keep First Column Of Array

*Wrap with formula to return first column of array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TAKE(\[\[ActiveCell.SpillParent::Formula\]\],,1)</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <code>kfc</code> |

[^Top](#oa-robot-definitions)

<BR>

### Keep First N Columns Of Array

*Wrap with formula to return first N columns of array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TAKE(\[\[ActiveCell.SpillParent::Formula\]\],,{{Number\_Of\_Columns}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>kfc</code></li><li><code>kfnc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Keep First N Rows Of Array

*Wrap with formula to return first N rows of array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TAKE(\[\[ActiveCell.SpillParent::Formula\]\],{{Number\_Of\_Rows}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>kfr</code></li><li><code>ktr</code></li><li><code>kfnr</code></li><li><code>ktnr</code></li><li><code>kn</code></li><li><code>kfn</code></li><li><code>ktn</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Keep First Row Of Array

*Wrap with formula to return first row of array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TAKE(\[\[ActiveCell.SpillParent::Formula\]\],1)</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>kfr</code></li><li><code>ktr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Keep Last Column Of Array

*Wrap with formula to return last column of array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TAKE(\[\[ActiveCell.SpillParent::Formula\]\],,\-1)</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <code>klc</code> |

[^Top](#oa-robot-definitions)

<BR>

### Keep Last N Columns Of Array

*Wrap with formula to return last N columns of array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TAKE(\[\[ActiveCell.SpillParent::Formula\]\],,\-{{Number\_Of\_Columns}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>klc</code></li><li><code>klnc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Keep Last N Rows Of Array

*Wrap with formula to return last N rows of array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TAKE(\[\[ActiveCell.SpillParent::Formula\]\],\-{{Number\_Of\_Rows}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>klr</code></li><li><code>kbr</code></li><li><code>klnr</code></li><li><code>kbnr</code></li><li><code>kbn</code></li><li><code>kln</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Keep Last Row Of Array

*Wrap with formula to return last row of array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TAKE(\[\[ActiveCell.SpillParent::Formula\]\],\-1)</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>klr</code></li><li><code>kbr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Keep Rows Of Array

*Wraps array formula with CHOOSECOLS of all selected rows.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=CHOOSEROWS(\[\[ActiveCell.SpillParent::Formula\]\],{{[Selected\_Row\_Indexes\_In\_Spilling\_Range](#selected_row_indexes_in_spilling_range)}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <code>kr</code> |

[^Top](#oa-robot-definitions)

<BR>

### Max Of Array

*Wrap with MAX function to return the maximum.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=MAX(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <code>max</code> |

[^Top](#oa-robot-definitions)

<BR>

### Max Of Array By Column

*Wrap with formula to return max of each column.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=BYCOL(\[\[ActiveCell::Formula\]\],LAMBDA(x,MAX(x)))</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>maxc</code></li><li><code>maxbc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Max Of Array By Row

*Wrap with formula to return max of each row.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=BYROW(\[\[ActiveCell::Formula\]\],LAMBDA(x,MAX(x)))</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>maxr</code></li><li><code>maxbr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Median Of Array

*Wrap with MEDIAN function to return the median.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=MEDIAN(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <code>med</code> |

[^Top](#oa-robot-definitions)

<BR>

### Min Of Array

*Wrap with MIN function to return the minimum.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=MIN(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <code>min</code> |

[^Top](#oa-robot-definitions)

<BR>

### Min Of Array By Column

*Wrap with formula to return min of each column.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=BYCOL(\[\[ActiveCell::Formula\]\],LAMBDA(x,MIN(x)))</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>minc</code></li><li><code>minbc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Min Of Array By Row

*Wrap with formula to return min of each row.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=BYROW(\[\[ActiveCell::Formula\]\],LAMBDA(x,MIN(x)))</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>minr</code></li><li><code>minbr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Paste Characters

*Paste copied cells split by character using SplitByCharacter lambda.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SplitByCharacter(\[\[Clipboard\]\])</code> |
| Formula Dependencies | <ol><li>[SplitByCharacter.lambda](#splitbycharacterlambda)</li><li>[TILE.lambda](#tilelambda)</li></ol> |
| User Context Filter | ExcelSelectionIsSingleCell AND ClipboardHasExcelData AND ExcelSelectionIsEmpty |
| Launch Codes | <code>pc</code> |

[^Top](#oa-robot-definitions)

<BR>

### Paste Count Unique Values In Array

*Paste Counts of each unique value in the copied range.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

[How to Use Count Unique Values In Array Lambda Command](<.\Documentation\Count Unique Values In Array.md>)

> \*\*Note:\*\* Great for using in competitions to answer questions like "which brand is mentioned the most in the data?"

| Property | Value |
| --- | --- |
| Formula | <code>\=CountEachUniqueValue(\[\[Clipboard\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | [CountEachUniqueValue.lambda](#counteachuniquevaluelambda) |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ClipboardHasExcelData AND ExcelActiveCellIsEmpty AND ExcelSelectionIsSingleCell |
| Documentation | [Count Unique Values In Array.md](#count-unique-values-in-arraymd) |
| Launch Codes | <code>pcu</code> |

[^Top](#oa-robot-definitions)

<BR>

### Paste Distinct\/Unique Row Values

*Pastes the values of distinct\/unique rows of the copied cells to the active cell.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=UNIQUE(\[\[Clipboard\]\])</code> |
| Convert To Values | Yes |
| User Context Filter | ClipboardHasExcelData AND ExcelActiveCellIsEmpty |
| Launch Codes | <ol><li><code>pd</code></li><li><code>pu</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Paste Lookup

*Wrap active cell with XLOOKUP function to lookup each value in first row of copied cells and return corresponding last column of copied cells.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=XLOOKUP(\[\[ActiveCell::Formula\]\],CHOOSECOLS(\[\[Clipboard::AddressAbsoluteAbsolute\]\],1),CHOOSECOLS(\[\[Clipboard::AddressAbsoluteAbsolute\]\],\-1))</code> |
| User Context Filter | ClipboardHasExcelData AND ExcelActiveCellIsNotEmpty |
| Launch Codes | <code>pl</code> |

[^Top](#oa-robot-definitions)

<BR>

### Paste Sequence Across\/Right

*Pastes a SEQUENCE function in the active cell spilling across to right based on the copied integer value.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Paste` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SEQUENCE(,\[\[Clipboard\]\])</code> |
| User Context Filter | ClipboardHasExcelData |
| Launch Codes | <ol><li><code>ps</code></li><li><code>psr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Paste Sequence Down

*Pastes a SEQUENCE function in the active cell spilling down based on the copied integer value.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Paste` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SEQUENCE(\[\[Clipboard\]\])</code> |
| User Context Filter | ClipboardHasExcelData |
| Launch Codes | <ol><li><code>ps</code></li><li><code>psd</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Paste Unpivot Array

*Pastes array formula with function to unpivot copied range by all column headers to the right.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=UNPIVOT(\[\[Clipboard\]\])</code> |
| Formula Dependencies | <ol><li>[UNPIVOT.lambda](#unpivotlambda)</li><li>[CROSSJOIN.lambda](#crossjoinlambda)</li></ol> |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ClipboardHasExcelData AND ExcelActiveCellIsEmpty AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>pu</code></li><li><code>pua</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Product Of Array

*Wrap with PRODUCT function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=PRODUCT(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent |

[^Top](#oa-robot-definitions)

<BR>

### Product Of Array By Column

*Wrap with formula to return product of each column.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=BYCOL(\[\[ActiveCell::Formula\]\],LAMBDA(x,PRODUCT(x)))</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>prod</code></li><li><code>prodc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Product Of Array By Row

*Wrap with formula to return product of each row.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=BYROW(\[\[ActiveCell::Formula\]\],LAMBDA(x,PRODUCT(x)))</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>prod</code></li><li><code>prodr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Blank Rows And Columns

*Wrap existing formula with RemoveBlanks Lambda function to remove blanks rows and columns from array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=RemoveBlanks(\[\[ActiveCell.SpillParent::Formula\]\])</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | [RemoveBlanks.lambda](#removeblankslambda) |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>rb</code></li><li><code>rbr</code></li><li><code>rbc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Columns Of Array

*Removes specified columns of array using RemoveCols lambda.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

\!\[OARobotImage\](oarobot:\/\/RemoveColumnsRowsOfArraypng)

| Property | Value |
| --- | --- |
| Formula | <code>\=RemoveCols(\[\[ActiveCell.SpillParent::Formula\]\], {{[Selected\_Column\_Indexes\_In\_Spilling\_Range](#selected_column_indexes_in_spilling_range)}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Formula Dependencies | [RemoveCols.lambda](#removecolslambda) |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <code>rc</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove First Column Of Array

*Wrap with DROP function to drop first column of array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=DROP(\[\[ActiveCell.SpillParent::Formula\]\],,1)</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <code>rfc</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove First N Columns Of Array

*Wrap with DROP function to drop first N columns of array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=DROP(\[\[ActiveCell.SpillParent::Formula\]\],,{{Number\_Of\_Columns}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>rfc</code></li><li><code>rfnc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Remove First N Rows Of Array

*Wrap with DROP function to drop first N rows of array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=DROP(\[\[ActiveCell.SpillParent::Formula\]\],{{Number\_Of\_Rows}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>rfr</code></li><li><code>rtr</code></li><li><code>rfn</code></li><li><code>rfnr</code></li><li><code>rtn</code></li><li><code>rtnr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Remove First Row Of Array

*Wrap with DROP function to drop first row of array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=DROP(\[\[ActiveCell.SpillParent::Formula\]\],1)</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>rfr</code></li><li><code>rtr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Last Column Of Array

*Wrap with DROP function to drop last column of array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=DROP(\[\[ActiveCell.SpillParent::Formula\]\],,\-1)</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <code>rlc</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Last N Columns Of Array

*Wrap with DROP function to drop last N columns of array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=DROP(\[\[ActiveCell.SpillParent::Formula\]\],,\-{{Number\_Of\_Columns}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>rlc</code></li><li><code>rlnc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Last N Rows Of Array

*Wrap with DROP function to drop last N rows of array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=DROP(\[\[ActiveCell.SpillParent::Formula\]\],\-{{Number\_Of\_Rows}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>rlr</code></li><li><code>rbr</code></li><li><code>rfn</code></li><li><code>rfnr</code></li><li><code>rbn</code></li><li><code>rbnr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Last Row Of Array

*Wrap with DROP function to remove last row of array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=DROP(\[\[ActiveCell.SpillParent::Formula\]\],\-1)</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>rlr</code></li><li><code>rbr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Other Columns Of Array

*Wraps array formula with CHOOSECOLS of all selected columns.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=CHOOSECOLS(\[\[ActiveCell.SpillParent::Formula\]\],{{[Selected\_Column\_Indexes\_In\_Spilling\_Range](#selected_column_indexes_in_spilling_range)}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <code>roc</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Other Rows Of Array

*Wraps array formula with CHOOSECOLS of all selected rows.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=CHOOSEROWS(\[\[ActiveCell.SpillParent::Formula\]\],{{[Selected\_Row\_Indexes\_In\_Spilling\_Range](#selected_row_indexes_in_spilling_range)}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <code>ror</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Rows Of Array

*Removes specified rows of array using RemoveRows lambda.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

\!\[OARobotImage\](oarobot:\/\/MyImagegif)

| Property | Value |
| --- | --- |
| Formula | <code>\=RemoveRows(\[\[ActiveCell.SpillParent::Formula\]\], {{[Selected\_Row\_Indexes\_In\_Spilling\_Range](#selected_row_indexes_in_spilling_range)}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Formula Dependencies | [RemoveRows.lambda](#removerowslambda) |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <code>rr</code> |

[^Top](#oa-robot-definitions)

<BR>

### Reshape To One (1) Column

*Wrap with TOCOL function to return as a column.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TOCOL(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <ol><li><code>1c</code></li><li><code>toc</code></li><li><code>tocol</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Reshape To One (1) Row

*Wrap with TOROW function to return as a row.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TOROW(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <ol><li><code>1r</code></li><li><code>tor</code></li><li><code>torow</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Reverse Columns (Flip Array Horizontally)

*Wrap existing formula with ReverseColumns Lambda function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=ReverseColumns(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | [ReverseColumns.lambda](#reversecolumnslambda) |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <ol><li><code>rev</code></li><li><code>revc</code></li><li><code>fa</code></li><li><code>fah</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Reverse Rows (Flip Array Vertically)

*Wrap existing formula with ReverseRows Lambda function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=ReverseRows(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | [ReverseRows.lambda](#reverserowslambda) |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <ol><li><code>rev</code></li><li><code>revr</code></li><li><code>fa</code></li><li><code>fav</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Running Product Of Array

*Wrap with RunningProduct Lambda function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=RunningProduct(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | [RunningProduct.lambda](#runningproductlambda) |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <code>rp</code> |

[^Top](#oa-robot-definitions)

<BR>

### Running Total Of Array

*Wrap with RunningTotal Lambda function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=RunningTotal(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | [RunningTotal.lambda](#runningtotallambda) |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <code>rt</code> |

[^Top](#oa-robot-definitions)

<BR>

### Running Totals By Column

*Wrap with RunningTotalsByColumn Lambda function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=RunningTotalsByColumn(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | [RunningTotalsByColumn.lambda](#runningtotalsbycolumnlambda) |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>rt</code></li><li><code>rtc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Running Totals By Row

*Wrap with RunningTotalsByRow Lambda function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Lambda` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=RunningTotalsByRow(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | [RunningTotalsByRow.lambda](#runningtotalsbyrowlambda) |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>rt</code></li><li><code>rtr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Sort Array Asc

*Wraps array formula with SORT function by selected columns of active array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SORT(\[\[ActiveCell.SpillParent::Formula\]\],{{[Selected\_Column\_Indexes\_In\_Spilling\_Range](#selected_column_indexes_in_spilling_range)}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <code>sa</code> |

[^Top](#oa-robot-definitions)

<BR>

### Sort Array Desc

*Wraps array formula with SORT function by selected columns of active array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SORT(\[\[ActiveCell.SpillParent::Formula\]\],{{[Selected\_Column\_Indexes\_In\_Spilling\_Range](#selected_column_indexes_in_spilling_range)}},\-1)</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <code>sad</code> |

[^Top](#oa-robot-definitions)

<BR>

### Split By Character

*Split active cell by character as row vector using SplitByCharacter lambda.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SplitByCharacter(\[\[ActiveCell::Formula\]\])</code> |
| Formula Dependencies | <ol><li>[SplitByCharacter.lambda](#splitbycharacterlambda)</li><li>[TILE.lambda](#tilelambda)</li></ol> |
| User Context Filter | ExcelActiveCellIsNotEmpty |
| Launch Codes | <code>sbc</code> |

[^Top](#oa-robot-definitions)

<BR>

### Split By Digit

*Split active cell by digits as row vector using SplitByDigits lambda.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SplitByDigit(\[\[ActiveCell::Formula\]\])</code> |
| Formula Dependencies | <ol><li>[SplitByDigit.lambda](#splitbydigitlambda)</li><li>[TILE.lambda](#tilelambda)</li></ol> |
| User Context Filter | ExcelActiveCellIsNotEmpty |
| Launch Codes | <code>sbd</code> |

[^Top](#oa-robot-definitions)

<BR>

### Split Numbers From Text To Columns

*Given a text value, split any numbers from text into row array of numbers.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SplitNumbersFromText(\[\[ActiveCell::Formula\]\])</code> |
| Formula Dependencies | <ol><li>[SplitNumbersFromText.lambda](#splitnumbersfromtextlambda)</li><li>[TILE.lambda](#tilelambda)</li></ol> |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <code>sn</code> |

[^Top](#oa-robot-definitions)

<BR>

### Sum Array

*Wrap with SUM function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SUM(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <code>sum</code> |

[^Top](#oa-robot-definitions)

<BR>

### Sum Array By Column

*Wrap with formula to return sum of each column.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=BYCOL(\[\[ActiveCell::Formula\]\],LAMBDA(x,SUM(x)))</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>sbc</code></li><li><code>sumc</code></li><li><code>sumbc</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Sum Array By Row

*Wrap with formula to return sum of each row.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=BYROW(\[\[ActiveCell::Formula\]\],LAMBDA(x,SUM(x)))</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>sbr</code></li><li><code>sumr</code></li><li><code>sumbr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Transpose Array

*Wrap with TRANSPOSE function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TRANSPOSE(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <ol><li><code>t</code></li><li><code>ta</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Unique Columns Of Array

*Wrap with UNIQUE function to return the unique column values.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=UNIQUE(\[\[ActiveCell::Formula\]\], TRUE)</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <code>ur</code> |

[^Top](#oa-robot-definitions)

<BR>

### Unique Rows Of Array

*Wrap with UNIQUE function to return the unique values.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#WrapWith` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=UNIQUE(\[\[ActiveCell::Formula\]\])</code> |
| Scroll To Destination | ☐Yes ☑No |
| User Context Filter | ExcelActiveCellIsSpillParent |
| Launch Codes | <code>ur</code> |

[^Top](#oa-robot-definitions)

<BR>

### Unpivot Array

*Wraps array formula with function to unpivot by all column headers to the right.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=UNPIVOT(\[\[ActiveCell::Formula\]\])</code> |
| Formula Dependencies | <ol><li>[UNPIVOT.lambda](#unpivotlambda)</li><li>[CROSSJOIN.lambda](#crossjoinlambda)</li></ol> |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>u</code></li><li><code>up</code></li><li><code>ua</code></li><li><code>upa</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Unpivot Array (By Selected Columns)

*Wraps array formula with function to unpivot by selected column headers.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Formula Command` `#Vol1`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=UNPIVOT(\[\[ActiveCell.SpillParent::Formula\]\],\[\[Selection::ValueArray\]\])</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Formula Dependencies | <ol><li>[UNPIVOT.lambda](#unpivotlambda)</li><li>[CROSSJOIN.lambda](#crossjoinlambda)</li></ol> |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellIsInSpillingToRange AND ExcelSelectionIsSingleArea AND ExcelSelectionIsSingleRow AND ExcelSelectionIsMultipleColumns |
| Launch Codes | <ol><li><code>u</code></li><li><code>up</code></li><li><code>ua</code></li><li><code>upa</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

## Parameter Definitions

<BR>

### Selected\_Column\_Indexes\_In\_Spilling\_Range

*List which columns of spilled array are included in selection.*

<sup>`@Array Robot Vol 1.xlsm` `!VBA Macro Parameter` </sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modRange.ListTargetColumnIndexesInSpillingRange](./VBA/modRange.bas#L4)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Selected\_Row\_Indexes\_In\_Spilling\_Range

*List which rows of spilled array are included in selection.*

<sup>`@Array Robot Vol 1.xlsm` `!VBA Macro Parameter` </sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modRange.ListTargetRowIndexesInSpillingRange](./VBA/modRange.bas#L82)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Unselected\_Column\_Indexes\_In\_Spilling\_Range

*List which columns of spilled array are not included in selection.*

<sup>`@Array Robot Vol 1.xlsm` `!VBA Macro Parameter` </sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modRange.ListNonTargetColumnIndexesInSpillingRange](./VBA/modRange.bas#L40)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Unselected\_Row\_Indexes\_In\_Spilling\_Range

*List which rows of spilled array are not included in selection.*

<sup>`@Array Robot Vol 1.xlsm` `!VBA Macro Parameter` </sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modRange.ListNonTargetRowIndexesInSpillingRange](./VBA/modRange.bas#L118)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

## Connection Definitions

<BR>

### Lambda Robot

*Lambda Robot command collection.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Workbook Connection` </sup>

| Property | Value |
| --- | --- |
| Connection String | <code>Lambda Robot.xlam</code> |

[^Top](#oa-robot-definitions)

<BR>

## Text Definitions

<BR>

### Count Unique Values In Array.md

*How to Use Count Unique Values In Array Lambda Command*

<sup>`@Array Robot Vol 1.xlsm` `!Default Text` </sup>

| Property | Value |
| --- | --- |
| Text | [Count Unique Values In Array.md](<./Text/Count Unique Values In Array.md>) |
| Value | <code>\# Count Unique Values In Array</code><br><code></code><br><code>\#\# Usage</code><br><code></code><br><code>\* one</code><br><code>\* two </code><br><code>\* three</code><br><code></code><br><code>\#\# Example</code><br><code>\!\[OARobotImage\](oarobot:\/\/Count\_Unique\_Example.png)</code> |
| Content Type | Markdown |
| Markdown Id | <code>CountUniqueValuesInArraymd</code> |

[^Top](#oa-robot-definitions)

<BR>

### CountEachCharacter.lambda

*Counts how many of each character exists across all values in array.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [CountEachCharacter.lambda](<./Text/CountEachCharacter.lambda.txt>) |
| Value | #Error: Failed to get value for property: [Value] Issue: [Exception has been thrown by the target of an invocation.] |
| Content Type | ExcelFormula |
| Location | <code>CountEachCharacter</code> |
| Markdown Id | <code>CountEachCharacterlambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### CountEachUniqueValue.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [CountEachUniqueValue.lambda](<./Text/CountEachUniqueValue.lambda.txt>) |
| Value | <code>\/\*Returns counts of each unique value in array. \*\/</code><br><code>CountEachUniqueValue \= LAMBDA(array,\[exclude\], LET(</code><br><code> \\\\LambdaName, "CountUniqueValues",</code><br><code> \\\\CommandName, "Count Unique Values",</code><br><code> \\\\Description, "Returns counts of each unique value in array.",</code><br><code> \_unique, UNIQUE(TOCOL(IF(array \= "", "", array))),</code><br><code> \_include, BYROW(\_unique \<\> TOROW(exclude), LAMBDA(x, PRODUCT(N(x)))),</code><br><code> \_filtered, IF(ISOMITTED(exclude), \_unique, FILTER(\_unique, \_include)),</code><br><code> \_countif, MAP(\_filtered, LAMBDA(x, SUM(\-\-(array \= x)))),</code><br><code> Result, SORTBY(HSTACK(\_filtered, \_countif), \_countif, \-1),</code><br><code> Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>CountEachUniqueValue</code> |
| Markdown Id | <code>CountEachUniqueValuelambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### CROSSJOIN.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [CROSSJOIN.lambda](<./Text/CROSSJOIN.lambda.txt>) |
| Value | <code>CrossJoin \= LAMBDA(table1,table2,\[has\_headers\], LET(</code><br><code> \_HasHeaders, IF(ISOMITTED(has\_headers),TRUE,has\_headers),</code><br><code> \_Data1, IF(\_HasHeaders,DROP(table1,1),table1),</code><br><code> \_Data2, IF(\_HasHeaders,DROP(table2,1),table2),</code><br><code> \_D1Rows, ROWS(\_Data1),</code><br><code> \_D1Cols, COLUMNS(\_Data1),</code><br><code> \_D2Rows, ROWS(\_Data2),</code><br><code> \_D2Cols, COLUMNS(\_Data2),</code><br><code> \_OuterJoinedData, MAKEARRAY(\_D1Rows \* \_D2Rows, \_D1Cols + \_D2Cols,LAMBDA(i,j,</code><br><code> IF(j \<\= \_D1Cols, INDEX(\_Data1, ROUNDUP(i \/ \_D2Rows, 0), j), INDEX(\_Data2, MOD(i \- 1, \_D2Rows)+1, j \- \_D1Cols)))),</code><br><code> \_WithHeader, IF(\_HasHeaders,VSTACK(HSTACK(TAKE(table1, 1), TAKE(table2, 1)), \_OuterJoinedData),\_OuterJoinedData),</code><br><code> \_WithHeader</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>CROSSJOIN</code> |
| Markdown Id | <code>CROSSJOINlambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### FillFromAbove.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [FillFromAbove.lambda](<./Text/FillFromAbove.lambda.txt>) |
| Value | <code>FillFromAbove \= LAMBDA(array,\[criteria\], LET(</code><br><code> \\\\LambdaName, "FillFromAbove",</code><br><code> fnCriteria, IF(ISOMITTED(criteria), LAMBDA(x, x \= ""), criteria),</code><br><code> \_cols, SEQUENCE(1, COLUMNS(array)),</code><br><code> \_rows, SEQUENCE(ROWS(array)),</code><br><code> fnIfBlank, LAMBDA(x, IF(x \= "", "", x)),</code><br><code> fnFill, LAMBDA(vector,</code><br><code> SCAN(fnIfBlank(INDEX(vector, 1, 1)), vector, LAMBDA(s,x, IF(fnCriteria(x), s, x)))</code><br><code> ),</code><br><code> Tile(\_cols, LAMBDA(n, fnFill(INDEX(array, \_rows, n))))</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>FillFromAbove</code> |
| Markdown Id | <code>FillFromAbovelambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### FillFromBelow.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [FillFromBelow.lambda](<./Text/FillFromBelow.lambda.txt>) |
| Value | <code>FillFromBelow \= LAMBDA(array,\[criteria\], LET(</code><br><code> \\\\LambdaName, "FillFromBelow",</code><br><code> fnCriteria, IF(ISOMITTED(criteria), LAMBDA(x, x \= ""), criteria),</code><br><code> \_cols, SEQUENCE(1, COLUMNS(array)),</code><br><code> \_rows, SEQUENCE(ROWS(array)),</code><br><code> fnIfBlank, LAMBDA(x, IF(x \= "", "", x)),</code><br><code> fnFill, LAMBDA(vector,</code><br><code> SCAN(fnIfBlank(INDEX(vector, 1, 1)), vector, LAMBDA(s,x, IF(fnCriteria(x), s, x)))</code><br><code> ),</code><br><code> fnReverse, LAMBDA(vector, SORTBY(vector, \_rows, \-1)),</code><br><code> Result, Tile(\_cols, LAMBDA(n, fnReverse(fnFill(fnReverse(CHOOSECOLS(array, n)))))),</code><br><code> Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>FillFromBelow</code> |
| Markdown Id | <code>FillFromBelowlambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### FillFromLeft.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [FillFromLeft.lambda](<./Text/FillFromLeft.lambda.txt>) |
| Value | <code>FillFromLeft \= LAMBDA(array,\[criteria\], LET(</code><br><code> \\\\LambdaName, "FillFromLeft",</code><br><code> fnCriteria, IF(ISOMITTED(criteria), LAMBDA(x, x \= ""), criteria),</code><br><code> \_cols, SEQUENCE(1, COLUMNS(array)),</code><br><code> \_rows, SEQUENCE(ROWS(array)),</code><br><code> fnIfBlank, LAMBDA(x, IF(x \= "", "", x)),</code><br><code> fnFill, LAMBDA(vector,</code><br><code> SCAN(fnIfBlank(INDEX(vector, 1, 1)), vector, LAMBDA(s,x, IF(fnCriteria(x), s, x)))</code><br><code> ),</code><br><code> Result, Tile(\_rows, LAMBDA(n, fnFill(CHOOSEROWS(array, n)))),</code><br><code> Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>FillFromLeft</code> |
| Markdown Id | <code>FillFromLeftlambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### FillFromRight.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [FillFromRight.lambda](<./Text/FillFromRight.lambda.txt>) |
| Value | <code>FillFromRight \= LAMBDA(array,\[criteria\], LET(</code><br><code> \\\\LambdaName, "FillFromRight",</code><br><code> fnCriteria, IF(ISOMITTED(criteria), LAMBDA(x, x \= ""), criteria),</code><br><code> \_cols, SEQUENCE(1, COLUMNS(array)),</code><br><code> \_rows, SEQUENCE(ROWS(array)),</code><br><code> fnIfBlank, LAMBDA(x, IF(x \= "", "", x)),</code><br><code> fnFill, LAMBDA(vector,</code><br><code> SCAN(fnIfBlank(INDEX(vector, 1, 1)), vector, LAMBDA(s,x, IF(fnCriteria(x), s, x)))</code><br><code> ),</code><br><code> fnReverse, LAMBDA(vector, SORTBY(vector, \_cols, \-1)),</code><br><code> Result, Tile(\_rows, LAMBDA(n, fnReverse(fnFill(fnReverse(CHOOSEROWS(array, n)))))),</code><br><code> Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>FillFromRight</code> |
| Markdown Id | <code>FillFromRightlambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### IFBLANK.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Default Text` </sup>

| Property | Value |
| --- | --- |
| Text | [IFBLANK.lambda](<./Text/IFBLANK.lambda.txt>) |
| Value | <code>IFBLANK \= LAMBDA(array,value\_if\_blank,MAP(array,LAMBDA(val,IF(ISBLANK(val),value\_if\_blank,val))));</code> |
| Content Type | ExcelFormula |
| Markdown Id | <code>IFBLANKlambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### IfText.lambda

*Definition of IfText lambda function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [IfText.lambda](<./Text/IfText.lambda.txt>) |
| Value | #Error: Failed to get value for property: [Value] Issue: [Exception has been thrown by the target of an invocation.] |
| Content Type | ExcelFormula |
| Location | <code>IfText</code> |
| Markdown Id | <code>IfTextlambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### IsOne.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [IsOne.lambda](<./Text/IsOne.lambda.txt>) |
| Value | <code>IsOne \= LAMBDA(array,LET(</code><br><code> \\\\LambdaName, "IsOne",</code><br><code>array\=1)</code><br><code>);</code> |
| Content Type | ExcelFormula |
| Location | <code>IsOne</code> |
| Markdown Id | <code>IsOnelambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### IsZero.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [IsZero.lambda](<./Text/IsZero.lambda.txt>) |
| Value | <code>IsZero \= LAMBDA(array,LET(</code><br><code> \\\\LambdaName, "IsZero",</code><br><code>((array\=0)\*(array\<\>""))\<\>0)</code><br><code>);</code> |
| Content Type | ExcelFormula |
| Location | <code>IsZero</code> |
| Markdown Id | <code>IsZerolambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### RemoveBlanks.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [RemoveBlanks.lambda](<./Text/RemoveBlanks.lambda.txt>) |
| Value | <code>\/\*Remove all blank rows and columns. (array \- array of values to evaluate for blank rows and columns)\*\/</code><br><code>RemoveBlanks \= LAMBDA(array,\[remove\_rows\],\[remove\_columns\], LET(</code><br><code> \\\\LambdaName, "REMOVEBLANKS",</code><br><code> \\\\CommandName, "Remove Blanks",</code><br><code> \\\\Description, "Remove all blank rows and columns.",</code><br><code> \\\\Parameters, {"array","array of values to evaluate for blank rows and columns"},</code><br><code> \\\\Source, "Excel Robot",</code><br><code> \_NonBlanks, (\-\-ISBLANK(array) + (array \= "")) \= 0,</code><br><code> \_NonBlankColumns, FILTER(</code><br><code> SEQUENCE(1, COLUMNS(\_NonBlanks)),</code><br><code> (BYCOL(\-\-\_NonBlanks, LAMBDA(x, SUM(x))) \<\> 0)</code><br><code> + IF(ISOMITTED(remove\_columns), 0, 1 \- remove\_columns)</code><br><code> ),</code><br><code> \_NonBlankRows, FILTER(</code><br><code> SEQUENCE(ROWS(\_NonBlanks)),</code><br><code> (BYROW(\-\-\_NonBlanks, LAMBDA(x, SUM(x))) \<\> 0) + IF(ISOMITTED(remove\_rows), 0, 1 \- remove\_rows)</code><br><code> ),</code><br><code> \_Result, CHOOSEROWS(CHOOSECOLS(IF(array \= "", "", array), \_NonBlankColumns), \_NonBlankRows),</code><br><code> \_Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>RemoveBlanks</code> |
| Markdown Id | <code>RemoveBlankslambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### RemoveCols.lambda

*Definition of RemoveCols lambda function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [RemoveCols.lambda](<./Text/RemoveCols.lambda.txt>) |
| Value | <code>\/\*Removes specified columns of array using RemoveCols lambda. \*\/</code><br><code>RemoveCols \= LAMBDA(array,column\_indexes, LET(</code><br><code> \\\\LambdaName, "RemoveCols",</code><br><code> \\\\CommandName, "Remove Columns Of Array",</code><br><code> \\\\Description, "Removes specified columns of array using RemoveCols lambda.",</code><br><code> \_Seq, SEQUENCE(COLUMNS(array)),</code><br><code> \_Keep, ISERROR(MATCH(\_Seq, TOROW(column\_indexes), 0)),</code><br><code> \_Included, FILTER(\_Seq, \_Keep, TRUE),</code><br><code> \_Result, CHOOSECOLS(array, \_Included),</code><br><code> \_Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>RemoveCols</code> |
| Markdown Id | <code>RemoveColslambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### RemoveRows.lambda

*Definition of RemoveRows lambda function.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [RemoveRows.lambda](<./Text/RemoveRows.lambda.txt>) |
| Value | <code>\/\*Removes specified rows of array using RemoveRows lambda. \*\/</code><br><code>RemoveRows \= LAMBDA(array,row\_indexes, LET(</code><br><code> \\\\LambdaName, "RemoveRows",</code><br><code> \\\\CommandName, "Remove Rows Of Array",</code><br><code> \\\\Description, "Removes specified rows of array using RemoveRows lambda.",</code><br><code> \_Seq, SEQUENCE(ROWS(array)),</code><br><code> \_Keep, ISERROR(MATCH(\_Seq, TOROW(row\_indexes), 0)),</code><br><code> \_Included, FILTER(\_Seq, \_Keep, TRUE),</code><br><code> \_Result, CHOOSEROWS(array, \_Included),</code><br><code> \_Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>RemoveRows</code> |
| Markdown Id | <code>RemoveRowslambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### ReplaceBlanksWithOnes.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [ReplaceBlanksWithOnes.lambda](<./Text/ReplaceBlanksWithOnes.lambda.txt>) |
| Value | #Error: Failed to get value for property: [Value] Issue: [Exception has been thrown by the target of an invocation.] |
| Content Type | ExcelFormula |
| Location | <code>ReplaceBlanksWithOnes</code> |
| Markdown Id | <code>ReplaceBlanksWithOneslambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### ReplaceBlanksWithZeros.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [ReplaceBlanksWithZeros.lambda](<./Text/ReplaceBlanksWithZeros.lambda.txt>) |
| Value | #Error: Failed to get value for property: [Value] Issue: [Exception has been thrown by the target of an invocation.] |
| Content Type | ExcelFormula |
| Location | <code>ReplaceBlanksWithZeros</code> |
| Markdown Id | <code>ReplaceBlanksWithZeroslambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### ReplaceZerosWithBlanks.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [ReplaceZerosWithBlanks.lambda](<./Text/ReplaceZerosWithBlanks.lambda.txt>) |
| Value | <code>ReplaceZerosWithBlanks \= LAMBDA(array,LET(</code><br><code> \\\\LambdaName, "ReplaceZerosWithBlanks",</code><br><code> \\\\CommandName, "Replace Zeros With Blanks",</code><br><code> \\\\Description, "Returns the passed array but with blanks where there were zeros.",</code><br><code>IF(array\=0,"",array))</code><br><code>);</code> |
| Content Type | ExcelFormula |
| Location | <code>ReplaceZerosWithBlanks</code> |
| Markdown Id | <code>ReplaceZerosWithBlankslambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### ReverseColumns.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [ReverseColumns.lambda](<./Text/ReverseColumns.lambda.txt>) |
| Value | <code>ReverseColumns \= LAMBDA(array,LET(</code><br><code> \\\\LambdaName, "ReverseColumns",</code><br><code> \\\\CommandName, "Reverse Columns",</code><br><code> \\\\Description, "Returns an array in reverse column order.",</code><br><code>SORTBY(array,SEQUENCE(1,COLUMNS(array)),\-1))</code><br><code>);</code> |
| Content Type | ExcelFormula |
| Location | <code>ReverseColumns</code> |
| Markdown Id | <code>ReverseColumnslambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### ReverseRows.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [ReverseRows.lambda](<./Text/ReverseRows.lambda.txt>) |
| Value | <code>ReverseRows \= LAMBDA(array,LET(</code><br><code> \\\\LambdaName, "ReverseRows",</code><br><code> \\\\CommandName, "Reverse Rows",</code><br><code> \\\\Description, "Returns array in reverse row order.",</code><br><code>SORTBY(array,SEQUENCE(ROWS(array)),\-1))</code><br><code>);</code> |
| Content Type | ExcelFormula |
| Location | <code>ReverseRows</code> |
| Markdown Id | <code>ReverseRowslambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### RunningProduct.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [RunningProduct.lambda](<./Text/RunningProduct.lambda.txt>) |
| Value | <code>RunningProduct \= LAMBDA(array,LET(</code><br><code> \\\\LambdaName, "RunningProduct",</code><br><code> \\\\CommandName, "Running Product",</code><br><code> \\\\Description, "Returns vector of running product.",</code><br><code> res, SCAN(1,array,LAMBDA(s,a,s\*a)),</code><br><code> res</code><br><code>) );</code> |
| Content Type | ExcelFormula |
| Location | <code>RunningProduct</code> |
| Markdown Id | <code>RunningProductlambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### RunningTotal.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [RunningTotal.lambda](<./Text/RunningTotal.lambda.txt>) |
| Value | <code>RunningTotal \= LAMBDA(array,LET( \\\\LambdaName, "RunningTotal",</code><br><code> res, SCAN(0,array,LAMBDA(s,a,s+a)),</code><br><code> res</code><br><code>) );</code> |
| Content Type | ExcelFormula |
| Location | <code>RunningTotal</code> |
| Markdown Id | <code>RunningTotallambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### RunningTotalsByColumn.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [RunningTotalsByColumn.lambda](<./Text/RunningTotalsByColumn.lambda.txt>) |
| Value | <code>RunningTotalsByColumn \= LAMBDA(array,LET( \\\\LambdaName, "RunningTotalsByColumn", Result, MAKEARRAY(ROWS(array),COLUMNS(array),LAMBDA(x,y,SUM(INDEX(array,SEQUENCE(x),y)))), Result ) );</code> |
| Content Type | ExcelFormula |
| Location | <code>RunningTotalsByColumn</code> |
| Markdown Id | <code>RunningTotalsByColumnlambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### RunningTotalsByRow.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [RunningTotalsByRow.lambda](<./Text/RunningTotalsByRow.lambda.txt>) |
| Value | <code>RunningTotalsByRow \= LAMBDA(array,LET( \\\\LambdaName, "RunningTotalsByRow",</code><br><code> Result, MAKEARRAY(ROWS(array),COLUMNS(array),LAMBDA(x,y,SUM(INDEX(array,x,SEQUENCE(1,y))))),</code><br><code> Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>RunningTotalsByRow</code> |
| Markdown Id | <code>RunningTotalsByRowlambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### SplitByCharacter.lambda

*Lambda to split text into row vector of characters.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [SplitByCharacter.lambda](<./Text/SplitByCharacter.lambda.txt>) |
| Value | <code>SplitByCharacter \= LAMBDA(text\_vector, LET(</code><br><code> \\\\LambdaName, "SplitByCharacter",</code><br><code> fnSingle, LAMBDA(text, LET(</code><br><code> \_Sequence, SEQUENCE(LEN(text)),</code><br><code> \_Transpose, COLUMNS(text\_vector) \= 1,</code><br><code> \_Split, MID(text, IF(\_Transpose, TRANSPOSE(\_Sequence), \_Sequence), 1),</code><br><code> \_Result, IF(LEN(text) \= 0, "", \_Split),</code><br><code> \_Result</code><br><code> )),</code><br><code> \_IsArray, TYPE(text\_vector) \= 64,</code><br><code> \_Result, IF(\_IsArray, IFNA(Tile(text\_vector, LAMBDA(x, fnSingle(x))), ""), fnSingle(text\_vector)),</code><br><code> \_Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>SplitByCharacter</code> |
| Markdown Id | <code>SplitByCharacterlambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### SplitByDigit.lambda

*Lambda to split text into row vector of characters.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [SplitByDigit.lambda](<./Text/SplitByDigit.lambda.txt>) |
| Value | <code>SplitByDigit \= LAMBDA(text\_vector, LET(</code><br><code> \\\\LambdaName, "SplitByDigits",</code><br><code> fnSingle, LAMBDA(text, LET(</code><br><code> \_Sequence, SEQUENCE(LEN(text)),</code><br><code> \_Transpose, COLUMNS(text\_vector) \= 1,</code><br><code> \_Split, MID(text, IF(\_Transpose, TRANSPOSE(\_Sequence), \_Sequence), 1),</code><br><code> \_Result, IF(LEN(text) \= 0, "", IFERROR(VALUE(\_Split), 0)),</code><br><code> \_Result</code><br><code> )),</code><br><code> \_IsArray, TYPE(text\_vector) \= 64,</code><br><code> \_Result, IF(\_IsArray, IFNA(Tile(text\_vector, LAMBDA(x, fnSingle(x))), ""), fnSingle(text\_vector)),</code><br><code> \_Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>SplitByDigit</code> |
| Markdown Id | <code>SplitByDigitlambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### SplitNumbersFromText.lambda

*Lambda to split numbers out of text as separate columns.*

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [SplitNumbersFromText.lambda](<./Text/SplitNumbersFromText.lambda.txt>) |
| Value | <code>SplitNumbersFromText \= LAMBDA(text\_vector,\[ignore\_commas\], LET(</code><br><code> \\\\LambdaName, "SplitNumbersFromText",</code><br><code> fnSplit, LAMBDA(text,\[ignore\_commas\],\[transpose\_result\], LET(</code><br><code> \_chars, MID(text, SEQUENCE(LEN(text)), 1),</code><br><code> \_isNumber, LET(</code><br><code> codes, CODE(\_chars),</code><br><code> priorCodes, VSTACK("", DROP(codes, \-1)),</code><br><code> nextCodes, VSTACK(DROP(codes, 1), ""),</code><br><code> fnIsDigit, LAMBDA(x, (x \>\= CODE("0")) \* (x \<\= CODE("9")) \= 1),</code><br><code> IsPeriod, codes \= CODE("."),</code><br><code> IsComma, (codes \= CODE(",")) \* ((ISOMITTED(ignore\_commas) + (ignore\_commas \= FALSE))) \<\> 0,</code><br><code> fnIsDigit(codes) + (IsPeriod + IsComma) \* fnIsDigit(priorCodes) \* fnIsDigit(nextCodes)</code><br><code> ),</code><br><code> \_flip, VSTACK(DROP(\_isNumber, 1) \<\> DROP(\_isNumber, \-1), FALSE),</code><br><code> \_addDelimiter, \_chars & IF(\_flip, "‡", ""),</code><br><code> \_split, TEXTSPLIT(CONCAT(\_addDelimiter), "‡"),</code><br><code> \_toNumbers, IFERROR(\-\-\_split, \_split),</code><br><code> \_result, IF(transpose\_result, TRANSPOSE(\_toNumbers), \_toNumbers),</code><br><code> \_result</code><br><code> )),</code><br><code> \_Transpose, COLUMNS(text\_vector) \<\> 1,</code><br><code> \_IsArray, TYPE(text\_vector) \= 64,</code><br><code> \_Result, IF(</code><br><code> \_IsArray,</code><br><code> IFNA(Tile(text\_vector, LAMBDA(x, fnSplit(x, ignore\_commas, \_Transpose))), ""),</code><br><code> fnSplit(text\_vector, ignore\_commas, FALSE)</code><br><code> ),</code><br><code> \_Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>SplitNumbersFromText</code> |
| Markdown Id | <code>SplitNumbersFromTextlambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### TILE.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [TILE.lambda](<./Text/TILE.lambda.txt>) |
| Value | <code>\/\*Tile the outputs of a single\-parameter function given an array map of parameters. (params \- array of parameters arranged how function results to be tiled; function \- single\-parameter Lambda name or function)\*\/</code><br><code>Tile \= LAMBDA(params,function,LET( \\\\LambdaName, "TILE", \\\\CommandName, "Tile", \\\\Description, "Tile the outputs of a single\-parameter function given an array map of parameters.", \\\\Parameters, {"params","array of parameters arranged how function results to be tiled";"function","single\-parameter Lambda name or function"}, \\\\Source, "Written by @ExcelRobot but inspired by Owen Price's STACKER lambda.",</code><br><code> firstrow, function(INDEX(params,1,1)),</code><br><code> stacker, LAMBDA(stack,param,VSTACK(stack,function(param))),</code><br><code> firstcol, IF(ROWS(params)\=1,firstrow,REDUCE(firstrow,DROP(TAKE(params,,1),1),stacker)),</code><br><code> Result, IF(COLUMNS(params)\=1,firstcol,HSTACK(firstcol,Tile(DROP(params,,1),function))),</code><br><code> Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>TILE</code> |
| Markdown Id | <code>TILElambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### UNPIVOT.lambda

<sup>`@Array Robot Vol 1.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [UNPIVOT.lambda](<./Text/UNPIVOT.lambda.txt>) |
| Value | <code>Unpivot \= LAMBDA(table,\[columns\_to\_unpivot\],\[attribute\_name\],\[value\_name\],\[remove\_blanks\], LET(</code><br><code> \_ColumnsToUnpivot, IF(</code><br><code> ISOMITTED(columns\_to\_unpivot),</code><br><code> DROP(TAKE(table, 1), , 1),</code><br><code> columns\_to\_unpivot</code><br><code> ),</code><br><code> \_AttributeLabel, IF(ISOMITTED(attribute\_name), "Attribute", attribute\_name),</code><br><code> \_ValueLabel, IF(ISOMITTED(value\_name), "Value", value\_name),</code><br><code> \_FirstColumnToUnpivot, MATCH(</code><br><code> INDEX(\_ColumnsToUnpivot, , 1),</code><br><code> INDEX(table, 1, ),</code><br><code> 0</code><br><code> ),</code><br><code> \_UnpivotColumnCount, COLUMNS(\_ColumnsToUnpivot),</code><br><code> \_ColumnNumbers, SEQUENCE(1, COLUMNS(table)),</code><br><code> \_IncludeColumns, (\_ColumnNumbers \>\= \_FirstColumnToUnpivot)</code><br><code> \* (\_ColumnNumbers \< \_FirstColumnToUnpivot + \_UnpivotColumnCount),</code><br><code> \_UnpivotColumns, FILTER(\_ColumnNumbers, \_IncludeColumns),</code><br><code> \_OtherColumns, FILTER(\_ColumnNumbers, NOT(\_IncludeColumns)),</code><br><code> \_FullOuterJoin, CrossJoin(</code><br><code> CHOOSECOLS(table, \_OtherColumns),</code><br><code> VSTACK(\_AttributeLabel, TRANSPOSE(\_ColumnsToUnpivot)),</code><br><code> TRUE</code><br><code> ),</code><br><code> \_WithValues, HSTACK(</code><br><code> \_FullOuterJoin,</code><br><code> VSTACK(\_ValueLabel, TOCOL(DROP(CHOOSECOLS(table, \_UnpivotColumns), 1)))</code><br><code> ),</code><br><code> \_RemoveBlanks, IF(</code><br><code> OR(ISOMITTED(remove\_blanks), remove\_blanks),</code><br><code> FILTER(\_WithValues, INDEX(\_WithValues, , COLUMNS(\_WithValues)) \<\> ""),</code><br><code> IF(\_WithValues \= "", "", \_WithValues)</code><br><code> ),</code><br><code> \_ColumnOrder, LET(</code><br><code> n, COLUMNS(\_RemoveBlanks),</code><br><code> s, SEQUENCE(1, n),</code><br><code> IFS(</code><br><code> s \< \_FirstColumnToUnpivot,</code><br><code> s,</code><br><code> s \< \_FirstColumnToUnpivot + 2,</code><br><code> s + n \- \_FirstColumnToUnpivot \- 1,</code><br><code> TRUE,</code><br><code> s \- 2</code><br><code> )</code><br><code> ),</code><br><code> \_Result, CHOOSECOLS(\_RemoveBlanks, \_ColumnOrder),</code><br><code> \_Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>UNPIVOT</code> |
| Markdown Id | <code>UNPIVOTlambda</code> |

[^Top](#oa-robot-definitions)

<BR>

## Image Definitions

<BR>

### Count\_Unique\_Example.png

<sup>`@Array Robot Vol 1.xlsm` `!Default Image` </sup>

| Property | Value |
| --- | --- |
| Value | ![OARobotImage](oarobot://CountUniqueExamplepng) |
| Markdown Id | <code>CountUniqueExamplepng</code> |

[^Top](#oa-robot-definitions)

<BR>

### MyImage.gif

<sup>`@Array Robot Vol 1.xlsm` `!Default Image` </sup>

| Property | Value |
| --- | --- |
| Value | ![OARobotImage](oarobot://MyImagegif) |
| Markdown Id | <code>MyImagegif</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove\_Columns\_Rows\_Of\_Array.png

*Remove Columns\/Rows Of Array*

<sup>`@Array Robot Vol 1.xlsm` `!Default Image` </sup>

| Property | Value |
| --- | --- |
| Value | ![OARobotImage](oarobot://RemoveColumnsRowsOfArraypng) |
| Markdown Id | <code>RemoveColumnsRowsOfArraypng</code> |

[^Top](#oa-robot-definitions)
