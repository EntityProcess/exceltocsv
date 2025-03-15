# ExcelToCsv

A command-line tool that converts Excel files (XLS/XLSX) to CSV format. This tool is specifically designed to work with Git's textconv feature, allowing you to search through Excel files in your Git history.

## Features

- Converts Excel files (XLS/XLSX) to CSV format
- Integrates with Git's textconv for searching through Excel file history
- Enables searching through Excel file contents using Git's search capabilities

## Build

### Prerequisites
- .NET SDK 

### Building from Source

1. Clone the repository
2. Open a terminal in the project directory
3. Run the following command to create a release build:

```bash
dotnet publish -c Release
```

The executable will be created in the following path:
```
bin/Release/net8.0/publish/ExcelToCsv.exe
```

You can either add this directory to your PATH or use the full path when configuring Git.

## Setup

### 1. Configure Git Attributes

Add the following lines to your `.gitattributes` file:

```
*.xls diff=excel
*.xlsx diff=excel
```

If you don't have a `.gitattributes` file, create one in your repository's root directory.

### 2. Configure Git's textconv

Run the following command to set up the Excel converter:

```bash
git config --global diff.excel.textconv "ExcelToCsv.exe"
```

## Usage

Once configured, you can search through your Excel files' content in Git history using:

```bash
git log -S"your search string" -- path/to/your/excel/files
```

For example:
```bash
git log -S"Revenue" -- reports/*.xlsx
```

This will search for the term "Revenue" in all Excel files in the reports directory throughout your Git history.

## How It Works

When Git needs to search through Excel files, it uses ExcelToCsv to convert the Excel files to a text format (CSV) behind the scenes. This allows Git to perform text-based searches on the contents of Excel files, which would otherwise be impossible due to their binary format.

## Notes

- The converter needs to be in your system PATH or you should provide the full path in the git config command
- This tool is particularly useful for tracking changes in Excel-based data files, configurations, or reports
- The conversion is done on-the-fly and doesn't modify your original Excel files