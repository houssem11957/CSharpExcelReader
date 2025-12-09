# Ultra-Fast Excel (.xlsx) to C# Entity Parser
A dependency-free, high-performance Excel reader for .NET that turns large `.xlsx` files into strongly typed C# entities in seconds.

---

## Overview

Parsing big Excel files in .NET shouldn’t take half an hour.

This project is a lightweight Excel parser that treats `.xlsx` files as what they really are: ZIP archives containing XML (OpenXML) files. By reading the XML directly, it avoids the heavy object models and memory overhead of traditional Excel libraries like EPPlus.

- Tested on ~200,000 rows × 8 columns (≈1.6M cells)
- EPPlus: 30+ minutes
- This parser: ~15 seconds
- No external dependencies

---

## Why This Exists

I had a real-world problem: importing large Excel files in a production system.

Using EPPlus (a great library, very flexible and easy to use) worked fine for small and medium files, but when the sheet reached hundreds of thousands of rows, performance collapsed. A single import could take over 30 minutes, and memory usage wasn’t pretty either.

I benchmarked multiple Excel libraries, tried different strategies and tuning options, but nothing gave me the speed and lightweight footprint I needed.

On top of that, company policy didn’t allow adding more heavy dependencies to the already large project.

So I took a different route:

- Change `.xlsx` → `.zip`
- Inspect the internal structure (`xl/`, `docProps/`, `_rels/`, etc.)
- Read sheets and shared strings at the XML level
- Build a small helper to convert rows into C# entities

That’s how this parser was born: a focused, fast helper that does one job very well — read Excel into entities, with minimal overhead.

---

## Features

- High-performance parsing for large `.xlsx` files
- No external dependencies (pure .NET)
- Low memory usage
- Maps rows directly to strongly typed C# entities
- Flexible column header → property name mapping
- Converts cell values to common .NET types:
  - `string`, `int`, `long`, `short`, `byte`
  - `decimal`, `double`, `float`
  - `bool`
  - `DateTime` (including Excel OLE dates)
  - `Guid`
- Graceful error handling (bad values are skipped, not fatal)

---

## Getting Started

### 1. Define your entity

```csharp
public class Person
{
    public int Id { get; set; }
    public string Name { get; set; }
    public DateTime DateOfBirth { get; set; }
    public string JobTitle { get; set; }
}
