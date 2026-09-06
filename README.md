# Excel Functions Reference Guide

A practical reference guide covering essential mathematical, statistical, and array functions in Microsoft Excel with clear explanations, syntax, and real-world business examples.

---

### 1. SUM

* **Explanation:** Adds all numeric values within a specified range of cells. Text and blank cells are automatically ignored.
* **Syntax:** `=SUM(number1, [number2], ...)`
* **Example Scenario:** Calculating the total revenue from individual daily store transactions.
* **Sample Data:**
* Range `B2:B5` contains: `1200`, `450`, `780`, `1100`


* **Formula:** `=SUM(B2:B5)`
* **Output:** `3530`

---

### 2. MIN

* **Explanation:** Identifies and returns the smallest (lowest) numeric value in a specified dataset.
* **Syntax:** `=MIN(number1, [number2], ...)`
* **Example Scenario:** Finding the shortest delivery transit time (in days) across a shipping fleet.
* **Sample Data:**
* Range `C2:C6` contains: `4`, `2`, `7`, `3`, `5`


* **Formula:** `=MIN(C2:C6)`
* **Output:** `2`

---

### 3. MAX

* **Explanation:** Returns the largest (highest) numeric value in a dataset.
* **Syntax:** `=MAX(number1, [number2], ...)`
* **Example Scenario:** Identifying the highest single sale order amount in a monthly report.
* **Sample Data:**
* Range `D2:D6` contains: `15000`, `22000`, `18500`, `31000`, `9400`


* **Formula:** `=MAX(D2:D6)`
* **Output:** `31000`

---

### 4. AVERAGE

* **Explanation:** Calculates the arithmetic mean of a selected range of numbers by dividing the sum of values by the total count of numeric cells.
* **Syntax:** `=AVERAGE(number1, [number2], ...)`
* **Example Scenario:** Evaluating the mean performance score of students in a semester exam.
* **Sample Data:**
* Range `E2:E5` contains: `80`, `70`, `90`, `60`


* **Formula:** `=AVERAGE(E2:E5)`
* **Output:** `75`

---

### 5. SUMPRODUCT

* **Explanation:** Multiplies corresponding elements in two or more arrays of the same size, then sums all calculated products without requiring a helper column.
* **Syntax:** `=SUMPRODUCT(array1, [array2], ...)`
* **Example Scenario:** Computing total invoice value by multiplying item quantities by their respective unit prices.
* **Sample Data:**
* Quantity sold (`A2:A4`): `5`, `10`, `2`
* Unit price (`B2:B4`): `20`, `15`, `50`


* **Formula:** `=SUMPRODUCT(A2:A4, B2:B4)`
* **Calculation:** `(5*20) + (10*15) + (2*50) = 100 + 150 + 100`
* **Output:** `350`

---

### 6. RAND

* **Explanation:** Generates an evenly distributed random floating-point decimal number greater than or equal to 0 and less than 1. The value recalculates automatically every time the worksheet updates.
* **Syntax:** `=RAND()`
* **Example Scenario:** Assigning random baseline probabilities for split testing (A/B testing) models.
* **Formula:** `=RAND()`
* **Output:** `0.7492...` (changes dynamically on recalculation)

---

### 7. RANDBETWEEN

* **Explanation:** Generates a random integer between two specified boundary values (`bottom` and `top`).
* **Syntax:** `=RANDBETWEEN(bottom, top)`
* **Example Scenario:** Generating synthetic data like mock 4-digit verification codes or dummy order IDs.
* **Formula:** `=RANDBETWEEN(1000, 9999)`
* **Output:** `4821`

---

### 8. MOD

* **Explanation:** Returns the remainder after a number is divided by a specified divisor.
* **Syntax:** `=MOD(number, divisor)`
* **Example Scenario:** Checking packaging inventory to see how many surplus items remain after packing products into boxes of 5.
* **Formula:** `=MOD(23, 5)`
* **Output:** `3`

---

### 9. ABS

* **Explanation:** Returns the absolute value of a given number, removing any negative sign and treating every value as non-negative.
* **Syntax:** `=ABS(number)`
* **Example Scenario:** Calculating budget-to-actual variance where the absolute size of the deviation matters rather than whether it is positive or negative.
* **Formula:** `=ABS(500 - 850)`
* **Output:** `350`

---

### 10. ROUND

* **Explanation:** Rounds a numeric value to a specified number of decimal places based on standard rounding rules (digits 5 and above round upward).
* **Syntax:** `=ROUND(number, num_digits)`
* **Example Scenario:** Standardizing financial transactions and sales tax amounts to two decimal places.
* **Formula:** `=ROUND(145.678, 2)`
* **Output:** `145.68`

---

### 11. PI

* **Explanation:** Returns the exact mathematical constant Pi ($3.14159265358979$) accurate to 15 digits. Takes no arguments.
* **Syntax:** `=PI()`
* **Example Scenario:** Calculating the base area of a cylindrical industrial storage tank ($Area = \pi \times r^2$).
* **Sample Data:**
* Radius in cell `A2`: `7`


* **Formula:** `=PI() * (A2^2)`
* **Output:** `153.94`

---

### 12. SORT

* **Explanation:** Dynamically sorts the contents of a range or array in ascending or descending order without altering the original source data.
* **Syntax:** `=SORT(array, [sort_index], [sort_order], [by_col])`
* **Example Scenario:** Sorting sales representatives by total revenue in descending order.
* **Sample Data:**
* Range `A2:B4` (Name & Sales): `Amit (40)`, `Pooja (85)`, `Rahul (60)`


* **Formula:** `=SORT(A2:B4, 2, -1)` *(Column 2, -1 for descending)*
* **Output:**
* `Pooja  85`
* `Rahul  60`
* `Amit   40`



---

### 13. LOG

* **Explanation:** Returns the logarithm of a number to a specified base. If the base parameter is omitted, it defaults to base 10.
* **Syntax:** `=LOG(number, [base])`
* **Example Scenario:** Applying log transformation to normalize heavily skewed financial datasets in exploratory data analysis.
* **Formula:** `=LOG(100, 10)`
* **Output:** `2`

---

### 14. COUNT

* **Explanation:** Counts the total number of cells in a range that contain numeric values (including integers, decimals, and dates). Text strings, errors, and empty cells are ignored.
* **Syntax:** `=COUNT(value1, [value2], ...)`
* **Example Scenario:** Determining the number of customer registration rows that include a valid numeric phone number.
* **Sample Data:**
* Range `F2:F6` contains: `9810123456`, `"N/A"`, `9876543210`, `""`, `8800112233`


* **Formula:** `=COUNT(F2:F6)`
* **Output:** `3`
