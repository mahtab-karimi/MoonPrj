If you test the LoanCalculator application with extreme data, you will discover that some additional error handling code may be required. Attempt to calculate the monthly payment for a loan of $1 paid off in 999,999 months with a reasonable interest rate. The result will be "NaN," which means that the calculations produced a value that's not a valid number. NaN stands for Not A Number, and this value is discussed in Chapter 3.
The monthly payment for the same amount paid off in 9,999 is zero! Actually, it's less than half a cent a month, which is rounded to 0.00.

Of course, the extreme data are not valid (in most cases), but even then your code should offer some help to the user. You can also impose limits to the values entered by the user: a minimum/maximum amount, a reasonable duration (from 1 month to 50 years) and a rasonable interest rate.

TAB ORDER
If you attempt to operate the LoanCalculator application with the keyboard, you may discover that the Tab key doesn't take you always to the next control on the form. You will learn how to set the Tab Order of the various controls on the form in Chapter 5. 