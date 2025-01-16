# BMEcat Exporter

BMEcat Exporter is a tool for processing BMEcat XML catalogs and exporting product data into CSV and Excel files with a hierarchical structure. This tool allows users to generate detailed files containing product information, features, and logistics details organized by categories.

## Features

- **CSV Export**: Generate CSV files with customizable columns for product features, logistics information, and category hierarchy.
- **Excel Export**: Generate Excel files (optional) for better data presentation.
- **Dynamic Columns**: Automatically detects and includes all product features and logistics details as separate columns in the exported files.
- **Category Hierarchy**: Organizes products by root, node, and leaf categories.

## Technologies Used

- **C#**: The core programming language.
- **EPPlus**: Library for generating Excel files.
- **.NET Core**: Framework for running the application.

## Prerequisites

1. Install .NET Core SDK:
   ```
   https://dotnet.microsoft.com/download
   ```
2. Install EPPlus library for Excel export:
   ```bash
   dotnet add package EPPlus --version 5.8.0
   ```
3. Add the BMEcat.net library to your project for processing BMEcat catalogs. This tool is based on the repository:
   [BMEcat.net by Stephan Stapel](https://github.com/stephanstapel/BMECat.net/tree/master).

   To add it:
   ```bash
   git clone https://github.com/stephanstapel/BMECat.net.git
   ```
   Follow the instructions in the repository to include the library in your project.

## Project Structure

```
.
├── Program.cs           # Main entry point of the application
├── ProductCatalog.cs    # Models for BMEcat catalog data
├── ExportFunctions.cs   # Functions for exporting data
├── Utils.cs             # Utility functions (e.g., file sanitization)
├── README.md            # Project documentation
```

## Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/bmecat-exporter.git
   ```
2. Navigate to the project directory:
   ```bash
   cd bmecat-exporter
   ```
3. Build the project:
   ```bash
   dotnet build
   ```
4. Run the application:
   ```bash
   dotnet run
   ```

## Usage

1. Place the BMEcat XML file (e.g., `picard_bmecat_de.xml`) in the project directory.
2. Open `Program.cs` and update the `filePath` and `outputDirectory` variables with the path to your BMEcat file and desired output location.
3. Run the project:
   ```bash
   dotnet run
   ```
4. The exported files will be saved in the specified output directory.

## Export Functions

### CSV Export

The `ExportToCsvWithHierarchy` function processes the BMEcat catalog and generates CSV files for each category.

### Excel Export (Optional)

The `ExportToExcelWithHierarchy` function generates Excel files with the same structure. Ensure EPPlus is installed before using this feature.

## Contributing

Contributions are welcome! If you find a bug or have an idea for a feature, please create an issue or submit a pull request.

## License

This project is licensed under the MIT License. See the LICENSE file for details.

## Contact

For any questions or feedback, feel free to reach out:

- **GitHub**: [pooyaPera](https://github.com/pooyaPera)
