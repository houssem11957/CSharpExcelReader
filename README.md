Below is a fully polished, beautifully structured, **professional-grade README.md** designed for a GitHub repository.
It is elegant, clear, impressive, and written to make your project look **top-tier**.


# ⚡ Ultra-Fast Excel (.xlsx) to C# Entity Parser

A blazing-fast, dependency-free Excel reader for .NET that treats `.xlsx` files as raw OpenXML (ZIP + XML). Designed to parse **massive spreadsheets** with minimal memory usage and transform them into strongly typed C# entities in seconds.


## 🚀 Why This Exists

Handling Excel files in large enterprise systems isn’t always fun — especially when traditional libraries take *forever*.

Importing a **200,000-row × 8-column** Excel sheet (~1.6 million cells) with EPPlus took **30+ minutes**, sometimes even longer. Great library, but not built for this scale.

This parser brings that down to **~15 seconds** by skipping the overhead and reading the Excel file at its core:

* XML data
* Shared strings
* Sheet relationships
* Type conversions

No dependencies. No heavyweight object models. No waiting half an hour.

Just clean, fast, predictable performance.


## ✨ Features

* **🚀 Extreme performance** — Parses huge `.xlsx` files in seconds
* **📦 Zero dependencies** — No EPPlus, ClosedXML, or Interop
* **🧠 Low memory usage** — Streams + direct XML parsing
* **📑 Strongly typed results** — Map rows into your C# entities
* **🔍 Flexible column mapping** — Match Excel headers to properties easily
* **🏢 Enterprise-friendly** — Perfect for environments with dependency restrictions
* **💥 Stable + predictable** — Ignore bad cells, process everything else



## 🧱 How It Works

Excel `.xlsx` files are actually ZIP archives.
Inside them are a bunch of XML files:

```
/xl/workbook.xml
/xl/worksheets/sheet1.xml
/xl/sharedStrings.xml
/xl/styles.xml
/_rels/workbook.xml.rels
/docProps/
```

This parser opens the file as a `ZipArchive` and directly reads the sheet and string data from XML.
No unnecessary layers. No large models. No delays.


## 🛠️ Example Usage

```csharp
var mapping = new ColumnMapping();
mapping.Add("Employee Id", "Id");
mapping.Add("Full Name", "Name");
mapping.Add("Date of birth", "DateOfBirth");
mapping.Add("Job Title", "JobTitle");

var people = ExcelReader.ReadToEntities<Person>(
    "employees.xlsx",
    mapping: mapping
);
```

**`people`** is now a fully populated `List<Person>`.


## 🧩 Column Mapping

Excel headers don’t always match your property names.
This takes care of that:

```csharp
mapping.Add("Name of the Person", "Name");
mapping.Add("Date of birth", "DateOfBirth");
mapping.Add("myId", "Id");
```

Mapping is case-insensitive and optional.


## 📉 Performance

| Rows    | Columns | Cells | EPPlus Time     | This Parser     |
| ------- | ------- | ----- | --------------- | --------------- |
| 200,000 | 8       | ~1.6M | **30+ minutes** | **~15 seconds** |

The secret?
Minimal memory usage + direct XML parsing + zero dependencies.


## 🧬 Example Entity

```csharp
public class Person
{
    public int Id { get; set; }
    public string Name { get; set; }
    public DateTime DateOfBirth { get; set; }
    public string JobTitle { get; set; }
}
```

## 🔍 What’s Inside the Code

* **ExcelReader.cs** → Core logic for reading sheets and cells
* **ColumnMapping.cs** → Maps Excel headers to entity properties
* **Person.cs** → Example model
* **Program.cs** → Usage demo

## 🧠 Ideal Use Cases

* ETL pipelines
* Enterprise tools with dependency restrictions
* Backend batch processors
* Import services handling huge Excel files
* Cloud-native workers that need predictable performance

## 🛡️ Requirements

* .NET 6 or later
* Reads `.xlsx` only (OpenXML format)

## 📜 License

MIT — free for personal & commercial use.



## ⭐ Support

If this saved you from another 30-minute Excel import, consider starring the repo!
Your support helps more developers discover faster ways to work with big data in .NET.


