# Microsoft-Excel-Project-8

# Case study

Jamie, at Adventure Works, is attending a meeting with a customer, Contoso Bikes. Help Jamie to complete a worksheet that contains a summary of the last order placed by Contoso Bikes. 

The worksheet needs the following additional information for Contoso Bikes:

⦁ Delivery charges

⦁ Discount rates

⦁ And regional totals
_____________________________________________________________________________________________________________________________________________________________________________________________________________________________________________

# Adding a Data Column Using the IFS Function in Excel

# Overview: 

In this exercise, I used logical functions in Microsoft Excel to create customized totals and calculate regional delivery charges for orders. The tasks involved using the IF, IFS, and SUMIF functions to perform calculations based on certain conditions.

# Key Tasks Completed:

# 1. Creating Discount Rates Based on Order Value:

     ⦁ Used the IF function to display a 10% discount if the subtotal exceeded $10,000, otherwise 0%.
     ⦁ Example Formula: =IF(G7>10000,10%,0)
     ⦁ Result: 10% (for a subtotal of $15,750).

# 2. Calculating Regional Delivery Charges Using the IFS Function:

     ⦁ Used the IFS function to determine delivery charges based on the region.
     ⦁ Example Formula: =IFS(J7="A",$D$2,J7="B",$D$3,J7="C",$D$4,TRUE,0)
     ⦁ Result: $75 (for Region B).

# 3. Adding the Total Excluding Delivery and Including Delivery Charge:

     ⦁ Created a formula to sum the total without the delivery amount and then added the calculated delivery charge.
     ⦁ Example Formula: =K7+L7
     ⦁ Result: $14,250.

# 4. Using SUMIF to Calculate Sales Totals by Region:

     ⦁ Used SUMIF to calculate sales totals for each region (A, B, and C).
     ⦁ Example Formula (for Region A): =SUMIF(J7:J16,"A",K7:K16)
     ⦁ Result: $40,382.50 (for Region A).

# Final Steps:

Autofill: Applied the double-click method to copy formulas down from row 7 to row 16.
Error Handling: Addressed formula inconsistency warnings by selecting "Ignore Error" where appropriate.

# Conclusion:

This exercise helped me practice and apply various logical functions in Excel to automate calculations for discounts, delivery charges, and regional sales totals. I now have a better understanding of how to use the IF, IFS, and SUMIF functions in real-world spreadsheet tasks.
