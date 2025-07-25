![Frame 1](https://github.com/user-attachments/assets/4ca99b3c-da89-42a8-8c7e-3158e9aefefb)
# Combining Multiple Excel Files Using Power Query (M Language)

This Power Query solution handles the challenge of combining multiple Excel files from a folder — each with **inconsistent formats**, **different sheet/tab names**, and **month headers in other languages**.

Some tables start on different rows or have non-standard headers, making this a tricky transformation task. This solution dynamically cleans and consolidates all files into one structured table.

```m
let
    // Used Folder.Files, this is automatic when you used Folder connector in Power Query.Also, this could be created inside a parameter.
    Source = Folder.Files(SourceofSalesByRegionData),

    //I considered filtering for .xlsx files only to not pick-up unrelated files.
    FilterExcelFiles =  Table.SelectRows(Source, each Text.EndsWith([Extension],".xlsx")),

    // I get content using Excel.WorkBook() function through Table.AddColumn().
    AddedCustom = Table.AddColumn(FilterExcelFiles, "Sales by Region", each Excel.Workbook([Content])),

    // I use list when removing other columns as part of best practice. 
    ColList = List.Select(
        Table.ColumnNames(AddedCustom),
        each _ = "Name" or _ = "Sales by Region"
    ),

    // Initial Cleaning - Part
    RemovedOtherColumns = Table.SelectColumns(AddedCustom, ColList),

    //Cleaning the names for each region.This could be a custom function.
    ExtractedTextBetweenDelimiters = 
        Table.TransformColumns(RemovedOtherColumns, 
        {
            {"Name",
            //I have nested let-in inside each row of the name to perform the functions below named "ExtractingtheName" and "Removinganyextensions" in the specified "Name" column.
               fxInitialCleaning,type text    
            }
        }           
    ),
   //This will delete prevent hidden tables.This could be a custom function.
     VisibleSheetsOnly = 
        Table.TransformColumns(ExtractedTextBetweenDelimiters, 
        {
            {"Sales by Region",
            //I have nested let-in inside each row of the name to perform the functions below named "Visible"
            fxVisibleSheets
            }
        }           
    ),
    //Changed the Name column inside the table in Sales by Region column into Category
    RenamedNestedTable = Table.TransformColumns(
    VisibleSheetsOnly,
    {
        {"Sales by Region", each Table.RenameColumns(_, {{"Name", "Category"}})}
    }
),
//We expand using the expand button using Powwer Query's UI.
   //We expand using the expand button using Powwer Query's UI.
    ExpandedSalesbyRegion = Table.ExpandTableColumn(
        VisibleSheetsOnly, 
        "Sales by Region", 
        {"Name", "Data", "Item", "Kind", "Hidden"}, 
        {"Category", "Data", "Item", "Kind", "Hidden"}),
//We challenged to inside the main table which can be located in the tables under the Data column.
    AddedYear = Table.AddColumn(
        ExpandedSalesbyRegion, 
        //we named the added column as "Year"
        "Year", 
        //for each row, we are navigating to the [Data] column from the current row which contains the main table, {3} gets the 4th row of the table as M is zero-based and [Column2] returns the value of the 4th row
        each [Data]{3}[Column2]
    ),
//This part contains all the necessary transformations needed for to produce a clean table. This could also be a function.
    TransformationsinTables = Table.AddColumn(AddedYear, "Custom", each fxConvertPortugeseToEnglish([Data])),

    Keepneededcolumns = Table.SelectColumns(TransformationsinTables,{"Custom","Name", "Category", "Year"}),
//This will be helpful in expanding dynamically especially when we switch our fxs from Emglish to Filipino and vise versa
   ColumnNamesToExpand = Table.ColumnNames(Keepneededcolumns{0}[Custom]),


    Expandedtablesinacolumn = Table.ExpandTableColumn(
        Keepneededcolumns, 
        "Custom", 
       ColumnNamesToExpand,
       ColumnNamesToExpand
    ),
    #"Changed Type" = Table.TransformColumnTypes(Expandedtablesinacolumn,{{"Product Line", type text}, {"Month", type text}, {"Sales", type number}, {"Name", type text}, {"Category", type text}, {"Year", Int64.Type}})
in
    #"Changed Type"
```
---

## 🛠️ What This Does

✔️ Connects to a folder of Excel files  
✔️ Dynamically identifies and processes multiple sheets  
✔️ Removes unnecessary rows and promotes headers  
✔️ Converts month names in different languages  
✔️ Combines everything into one clean, ready-to-use table

---

## 📦 Files Included

- `CombineExcelFiles.pq` – the full M code script  
- `workflow.png` – visual breakdown of the Power Query steps *(optional)*  
- `sample-files/` – sample folder structure or dummy Excel files *(if applicable)*

---

## 📚 Techniques Used

- Record, Table, and List manipulation  
- `Table.Skip`, `Table.PromoteHeaders`, and dynamic filtering  
- Folder and Sheet iteration  
- Applied Steps broken into reusable blocks

---

## 🙌 Credits

- Based on ideas and techniques from  
  📘 *Power Query: Beyond the User Interface* by **Chandeep Chhabra**  
  💡 Tips and guidance from **Pedro Bagtas**, senior M ninja 🥷

---

## 📎 LinkedIn Post

You can view the original post and visual walkthrough here:  
🔗 []

---

## 📬 Questions?

Feel free to connect or open an issue if you have questions or want to collaborate!
