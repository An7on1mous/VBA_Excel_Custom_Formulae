# VBA_Excel_Custom_Formulae
Create a custom formula or function using VBA.  
Sometimes at work we need to calculate specific metrics, KPIs or just simple percentages, with this principles you can create virtually any calculation you need for work. 

Having these calculations as custom functions can save some time, and storing them in a library in the company’s network makes it shareable to everyone and very easy to maintain.

About Functions

We also know them colloquially as formulas. However, a Formula is an equation designed by a user in Excel, while a Function is a predefined calculation in the spreadsheet application.

For example: =2+2 is a formula.

Whereas =SUM(A2,A3) is a function.

We are going to create our own functions just as we would use SUM, COUNT, VLOOKUP, etc.

To demonstrate how to do this, we are going to create a Gross Margin function, and a Gross Margin Percentage function.


Instructions

You can achieve this by following these steps:

Step 1: Open Visual Basic by pressing alt + F8 or going to Developer, Visual Basic.

Step 2: Add a new module if there is not one already. 

Step 3: Enter or copy and paste the following code, we are going to add the 2 functions in one step:

    Public Function CWGM(RetailSales As Long, RetailSalesAtCost As Long) As Double
      '	Calculates gross margin
      On Error Resume Next
        CWGM = RetailSales - RetailSalesAtCost
      On Error GoTo 0
    End Function
    
    Public Function CWGMPercentage(GMDollars As Long, Sales As Long) As Double
      '	Calculates gross margin percentage 
      On Error Resume Next
        If GMDollars = 0 Or Sales = 0 Then
          CWGMPercentage = 0
        Else
          CWGMPercentage = GMDollars / Sales
        End If
      On Error GoTo 0
    End Function

Step 4: Close Visual Basic

Step 5: In Excel, cell A1, enter this sample table:

Retail Sales	Retail Sales At Cost	Gross Margin $	Gross Margin %
1500	1000		
2750	2290		
1200	900		

Step 6:
On cell C2, type =CW and select CWGM, then select cells A2 for Retail Sales, and B2 for Retail Sales at cost, then copy the function down to C4.

Step 7:
On cell D2, type =CW and select CWGMPercentage, then select cells C2 for GMDollars,  and A2 for Retail Sales, then copy the function down to D4.

Step 8: To save the file, save it as .xlsm or better, .xlsb, since it has a VBA module in it.


How the Functions Work

Creating a Public Function in VBA makes it callable in Excel cells. 

Let’s take Gross Margin for example.
=CWGM(A2,B2)
After the name CWGM, in brackets, we declare the 2 parameters to be used in the function, which are Retail Sales and Retail Sales at Cost.

The VBA code:

    Public Function CWGM(RetailSales As Long, RetailSalesAtCost As Long) As Double
      On Error Resume Next
        CWGM = RetailSales - RetailSalesAtCost
      On Error GoTo 0
    End Function

We tell VBA that On Error, resume next, which means to ignore any errors and keep going without crashing. 
After this check, the real formula performs the calculation, which is retail sales MINUS sales at cost


In the case of CWGMPercentage
=CWGMPercentage(C2,A2)

    Public Function CWGMPercentage(GMDollars As Long, Sales As Long) As Double
      '	Calculates gross margin percentage 
      On Error Resume Next
        If GMDollars = 0 Or Sales = 0 Then
          CWGMPercentage = 0
        Else
          CWGMPercentage = GMDollars / Sales
        End If
      On Error GoTo 0
    End Function

We follow the same error handling procedure as before, but this time, we add an IF, THEN, ELSE statement to check if the divisor is zero, which would also return another error.

Thanks for reading, please follow and subscribe, and remember to visit www.sosa.tv for information about some data analytics projects I’ve worked on and other cool stuff.
