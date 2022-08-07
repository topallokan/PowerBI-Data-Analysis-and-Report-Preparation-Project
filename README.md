# PowerBI-Data-Analysis-and-Report-Preparation-Project
Business intelligence application on sample dataset

# Dataset
Northwind Dataset
The Northwind database contains the sales data for a fictitious company called Northwind Traders, which imports and exports specialty foods from around the world.
In this study, it was studied with the excel file obtained using the northwind dataset data.
 
# Get Data
Power Bi desktop opens and click on get data
![Get Data](https://user-images.githubusercontent.com/79374662/183286516-797cb357-4768-4e30-b288-8218adba5874.png)

Then click on the source from which you want to get data. We will use excel file in our work.
![Excel Workbook](https://user-images.githubusercontent.com/79374662/183286546-90fa26ff-662a-4a6e-b2c2-c9f96401692a.png)

and we choose the tables we want.
![Select Tables](https://user-images.githubusercontent.com/79374662/183286636-6b98701b-8160-466e-acde-8c95ec0e30b3.PNG)

# Data Preprocessing
The Query editor opens and previews the data. Tables are made to be related.
for example use first row as header in this table field click will solve the problem.
![Use First Head Rows](https://user-images.githubusercontent.com/79374662/183286918-91675893-2313-42a4-8793-25baf3892790.png)

The changes we make in the Query editor create a code in the M language on the back. This code allows us to structure our data. I will share M codes for changes made to other tables

**Categories**
```
let
    Source = Excel.Workbook(File.Contents("Your File Location\Data.xlsx"), null, true),
    Categories_Sheet = Source{[Item="Categories",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Categories_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"CategoryID", Int64.Type}, {"CategoryName", type text}, {"Description", type text}})
in
    #"Changed Type"
```
**Customers**
```
let
    Source = Excel.Workbook(File.Contents("Your File Location\Data.xlsx"), null, true),
    Customers_Sheet = Source{[Item="Customers",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Customers_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"CustomerID", type text}, {"CompanyName", type text}, {"ContactName", type text}, {"ContactTitle", type text}, {"Address", type text}, {"City", type text}, {"Region", type text}, {"PostalCode", type any}, {"Country", type text}, {"Phone", type text}, {"Fax", type text}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Region", "Phone", "Fax"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns", {"CustomerID"})
in
    #"Removed Duplicates"
```

**Employees**
```
let
    Source = Excel.Workbook(File.Contents("C:\Users\okan.topal\Desktop\Data.xlsx"), null, true),
    Employees_Sheet = Source{[Item="Employees",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Employees_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"EmployeeID", Int64.Type}, {"LastName", type text}, {"FirstName", type text}, {"Title", type text}, {"TitleOfCourtesy", type text}, {"BirthDate", type datetime}, {"HireDate", type datetime}, {"City", type text}, {"Region", type text}, {"PostalCode", type any}, {"Country", type text}, {"HomePhone", type text}, {"Extension", Int64.Type}, {"ReportsTo", type any}}),
    #"Removed Duplicates" = Table.Distinct(#"Changed Type", {"EmployeeID"})
in
    #"Removed Duplicates"
```

**Employee Territories**
```
let
    Source = Excel.Workbook(File.Contents("C:\Users\okan.topal\Desktop\Data.xlsx"), null, true),
    Employees_Sheet = Source{[Item="Employees",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Employees_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"EmployeeID", Int64.Type}, {"LastName", type text}, {"FirstName", type text}, {"Title", type text}, {"TitleOfCourtesy", type text}, {"BirthDate", type datetime}, {"HireDate", type datetime}, {"City", type text}, {"Region", type text}, {"PostalCode", type any}, {"Country", type text}, {"HomePhone", type text}, {"Extension", Int64.Type}, {"ReportsTo", type any}}),
    #"Removed Duplicates" = Table.Distinct(#"Changed Type", {"EmployeeID"})
in
    #"Removed Duplicates"
```

**Order Details**
```
let
    Source = Excel.Workbook(File.Contents("C:\Users\okan.topal\Desktop\Data.xlsx"), null, true),
    #"Order Details_Sheet" = Source{[Item="Order Details",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"Order Details_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"OrderID", Int64.Type}, {"ProductID", Int64.Type}, {"UnitPrice", type number}, {"Quantity", Int64.Type}, {"Discount", Percentage.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "Amount", each [UnitPrice] * [Quantity]),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "Discounted Amount", each [Amount] * [Discount]),
    #"Added Custom2" = Table.AddColumn(#"Added Custom1", "Net Amount", each [Amount] - [Discounted Amount]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Custom2",{{"Amount", Int64.Type}, {"Discounted Amount", Int64.Type}, {"Net Amount", Int64.Type}})
in
    #"Changed Type1"
```

**Orders**
```
let
    Source = Excel.Workbook(File.Contents("C:\Users\okan.topal\Desktop\Data.xlsx"), null, true),
    Orders_Sheet = Source{[Item="Orders",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Orders_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"OrderID", Int64.Type}, {"CustomerID", type text}, {"EmployeeID", Int64.Type}, {"OrderDate", type datetime}, {"ShipVia", Int64.Type}, {"Freight", type number}, {"ShipName", type text}, {"ShipAddress", type text}, {"ShipCity", type text}, {"ShipRegion", type text}, {"ShipPostalCode", type any}, {"ShipCountry", type text}})
in
    #"Changed Type"
```

**Products**
```
let
    Source = Excel.Workbook(File.Contents("C:\Users\okan.topal\Desktop\Data.xlsx"), null, true),
    Products_Sheet = Source{[Item="Products",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Products_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"ProductID", Int64.Type}, {"ProductName", type text}, {"SupplierID", Int64.Type}, {"CategoryID", Int64.Type}, {"QuantityPerUnit", type text}, {"UnitPrice", type number}, {"UnitsInStock", Int64.Type}, {"UnitsOnOrder", Int64.Type}, {"ReorderLevel", Int64.Type}, {"Discontinued", Int64.Type}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type", {"CategoryID"}, Categories, {"CategoryID"}, "Categories", JoinKind.LeftOuter),
    #"Expanded Categories" = Table.ExpandTableColumn(#"Merged Queries", "Categories", {"CategoryName", "Description"}, {"Categories.CategoryName", "Categories.Description"})
in
    #"Expanded Categories"
```


**Region**
```
let
    Source = Excel.Workbook(File.Contents("C:\Users\okan.topal\Desktop\Data.xlsx"), null, true),
    Region_Sheet = Source{[Item="Region",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Region_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"RegionID", Int64.Type}, {"RegionDescription", type text}})
in
    #"Changed Type"
```

**Shippers**
```
let
    Source = Excel.Workbook(File.Contents("C:\Users\okan.topal\Desktop\Data.xlsx"), null, true),
    Shippers_Sheet = Source{[Item="Shippers",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Shippers_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"ShipperID", Int64.Type}, {"CompanyName", type text}, {"Phone", type text}})
in
    #"Changed Type"
```

**Ships**
```
let
    Source = Excel.Workbook(File.Contents("C:\Users\okan.topal\Desktop\Data.xlsx"), null, true),
    Ships_Sheet = Source{[Item="Ships",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Ships_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"OrderID", Int64.Type}, {"ShippedDate", type datetime}, {"CustomerID", type text}, {"EmployeeID", Int64.Type}, {"ShipVia", Int64.Type}})
in
    #"Changed Type"
```

**Suppliers**
```
let
    Source = Excel.Workbook(File.Contents("C:\Users\okan.topal\Desktop\Data.xlsx"), null, true),
    Suppliers_Sheet = Source{[Item="Suppliers",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Suppliers_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"SupplierID", Int64.Type}, {"CompanyName", type text}, {"ContactName", type text}, {"ContactTitle", type text}, {"Address", type text}, {"City", type text}, {"Region", type text}, {"PostalCode", type any}, {"Country", type text}, {"Phone", type any}, {"Fax", type any}, {"HomePage", type text}})
in
    #"Changed Type"
```

**Suppliers**
```
let
    Source = Excel.Workbook(File.Contents("C:\Users\okan.topal\Desktop\Data.xlsx"), null, true),
    Suppliers_Sheet = Source{[Item="Suppliers",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Suppliers_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"SupplierID", Int64.Type}, {"CompanyName", type text}, {"ContactName", type text}, {"ContactTitle", type text}, {"Address", type text}, {"City", type text}, {"Region", type text}, {"PostalCode", type any}, {"Country", type text}, {"Phone", type any}, {"Fax", type any}, {"HomePage", type text}})
in
    #"Changed Type"
```

**Territories**
```
let
    Source = Excel.Workbook(File.Contents("C:\Users\okan.topal\Desktop\Data.xlsx"), null, true),
    Territories_Sheet = Source{[Item="Territories",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Territories_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"TerritoryID", Int64.Type}, {"TerritoryDescription", type text}, {"RegionID", Int64.Type}})
in
    #"Changed Type"
```

**Territories**
```
let
    Source = Excel.Workbook(File.Contents("C:\Users\okan.topal\Desktop\Data.xlsx"), null, true),
    Territories_Sheet = Source{[Item="Territories",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Territories_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"TerritoryID", Int64.Type}, {"TerritoryDescription", type text}, {"RegionID", Int64.Type}})
in
    #"Changed Type"
```

# Build Data Model
Using one-to-many logic, a relationship is established between fact and dimension tables.
![Model](https://user-images.githubusercontent.com/79374662/183288190-bc77c4ff-d7a5-4437-9df8-ec5e9fe68903.PNG)

# Report Preparing

As a result of all these processes, a file was obtained in which a report could be prepared.
Dashboards related to customers and their orders were prepared.

**Measures, Calculated Column and Calculated Tables in the report**

**DateTable**
```
Date = 
VAR MinYear = YEAR ( MIN ( Orders[OrderDate] ) )
VAR MaxYear = YEAR ( MAX ( Orders[OrderDate] ) )
RETURN
ADDCOLUMNS (
    FILTER (
        CALENDARAUTO( ),
        AND ( YEAR ( [Date] ) >= MinYear, YEAR ( [Date] ) <= MaxYear )
    ),
    "Calendar Year", "CY " & YEAR ( [Date] ),
    "Month Name", FORMAT ( [Date], "mmmm" ),
    "Month Number", MONTH ( [Date] ),
    "Weekday", FORMAT ( [Date], "dddd" ),
    "Weekday number", WEEKDAY( [Date] ),
    "Quarter", "Q" & TRUNC ( ( MONTH ( [Date] ) - 1 ) / 3 ) + 1
)
```

**Segment Table**
```
Segment Table = SUMMARIZE(Orders,Orders[OrderID],"Total Amount", 'Measure Table'[Total Amount])
)
```
**New Column**

 ```
 Priority = SWITCH(TRUE(),
 'Segment Table'[Total Amount]> 5000, "High",
'Segment Table'[Total Amount] > 1000, "Average",
'Segment Table'[Total Amount]  < 1000, "Low")
)
```

**Measure Table**
 ```
Average Delivery Time = var a = AVERAGE(Orders[Delivery Time])
var b = SUMMARIZE(Customers,Customers[ContactName],"delivery time", a)

return AVERAGEX(b,[Delivery time])
```

 ```
Customers Quantity = DISTINCTCOUNT(Customers[CustomerID])
```

```
Delivery time = DATEDIFF( SUM(Orders[OrderDate]),SUM(Ships[ShippedDate]),DAY)
```

```
Orders = DISTINCTCOUNT('Order Details'[OrderID])
```

```
Total Amount = SUM('Order Details'[Net Amount])
```

```
Total Amount - Freight = 'Measure Table'[Total Amount] - SUM(Orders[Freight])
```

```
Total Amount - Freight Per Order = DIVIDE( [Total Amount - Freight] , [Orders],BLANK())
```

```
Total Order Quantity = COUNT(Orders[OrderID])
```

**Note :**  Performance increase has been achieved by performing these 3 steps in the power query editor. The earlier the intervention is made at the source of the data, the more beneficial it will be for the report.
    
    #"Added Custom" = Table.AddColumn(#"Changed Type", "Amount", each [UnitPrice] * [Quantity]),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "Discounted Amount", each [Amount] * [Discount]),
    #"Added Custom2" = Table.AddColumn(#"Added Custom1", "Net Amount", each [Amount] - [Discounted Amount]),
















