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













