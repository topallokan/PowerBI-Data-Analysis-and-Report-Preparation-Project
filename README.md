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
    Source = Excel.Workbook(File.Contents("C:\Users\okan.topal\Desktop\Data.xlsx"), null, true),
    Categories_Sheet = Source{[Item="Categories",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Categories_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"CategoryID", Int64.Type}, {"CategoryName", type text}, {"Description", type text}})
in
    #"Changed Type"
```







