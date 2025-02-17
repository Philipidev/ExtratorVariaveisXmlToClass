# Word Template Parser

A simple C# console application that extracts tag values from a Word document and dynamically generates a corresponding C# class with properties for each tag.

## Overview

This project demonstrates how to use the [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/) library to read a Word document (DOCX) template, extract all `<w:tag w:val="..."/>` values, and generate a C# class where each property represents a tag value.

## Features

- Reads a Word document template.
- Extracts all tag values from `<w:tag w:val="..."/>` elements.
- Dynamically generates a C# class with properties matching the tag names.
- Utilizes the DocumentFormat.OpenXml library for Word document manipulation.

## Prerequisites

- [.NET SDK](https://dotnet.microsoft.com/download) (version 6.0 or later recommended)
- [Visual Studio](https://visualstudio.microsoft.com/) or any other C# IDE
- NuGet Package: [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/)

## Getting Started

### 1. Clone the Repository

```bash
git clone <repository-url>
```

### 2. Open the Project

Open the solution in Visual Studio or your preferred IDE.

### 3. Install the Required NuGet Package

Ensure that the DocumentFormat.OpenXml package is installed. You can install it via the Package Manager Console:

```powershell
Install-Package DocumentFormat.OpenXml
```

Or via the NuGet Package Manager UI.

### 4. Configure the Word Document Path

In `Program.cs`, update the `caminhoWord` variable with the full path to your Word document template (DOCX):

```csharp
string caminhoWord = @"C:\Path\To\Template_Copel.docx";
```

### 5. Build and Run

Build the project and run the console application. The generated C# class code will be printed to the console, similar to:

```csharp
public class Template_Copel
{
    public string DESCRICAO_ATUAL_ITEM { get; set; }
    // Additional properties based on the extracted tags...
}
```

## Project Structure

- **Program.cs:** Contains the main logic for extracting tag values and generating the C# class code.
- **README.md:** Provides documentation and instructions for using the application.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request for any improvements or bug fixes.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.