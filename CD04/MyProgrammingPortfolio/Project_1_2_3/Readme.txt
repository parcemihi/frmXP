Version 2 starts expanding the concept(s) for VB2 Project_1_2_3


Format 	ID - Cost - Notes - ... 


Databases:
   Products - Product ID, Product Cost, Notes, Product Name, Shipping Cost, Item Condition
   Sales    - Sales_ID  , Sales Cost  , Notes, Item         , Quantity      , Sale Flag, Flag Status
	Sales_ID = 

	EXPENSE_DRAWERSHIFT = 0
	EXPENSE_PAYOUT = 1
	EXPENSE_EMPLOYEE = 2
   Expenses - Expense ID, Expense Cost, Notes, Expense Type
   Customer
   Delivery
   
   

Processes: Filter
Invoice Number, Time, Cost, Quantity, Inventory, Product ID





 יא A program for cashier duties
 יא A database tracking product: sales/inventory
 יא To total the amount of a purchase
 יא Calulate and Add tax



Modules:
  Call_Inventory
   -Load List
   -Load Product ID

  -------------------------------------------------------------------------------------
  Call_Sales
   -Add new sale to database
   -Retrieve sales
     -All
     -FilterTime(StartTime,StopTime,Month,day)	'Default 12:00 am to 11:59 pm
     -FilterCost(FilterCostL, FilterCostH)	'Default 0 to -1
     -FilterQty(FilterQtyL, FilterQtyH)		'Default 0 to -1
     -



	----------------------------------------------------------------------------
	-	Filter_HiLo
	-	
	-
	-
	-
	-
  -------------------------------------------------------------------------------------

  Process_Sale
  Process_Tax
  
  
  
  
  PaymentOption
   -Cash
   -Credit
   -Check
  
  

  
  
  
  
  