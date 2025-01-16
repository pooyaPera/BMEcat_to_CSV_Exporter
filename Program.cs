using BMECat.net;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

class Program
{
    static void Main(string[] args)
    {
        // Der Pfad zur XML-Datei, die den Produktkatalog enthält
        string filePath = "";

        // Das Verzeichnis, in dem die CSV-Dateien exportiert werden
        string outputDirectory = "";


        // Laden des Produktkatalogs aus der XML-Datei
        ProductCatalog catalog = ProductCatalog.Load(filePath);

        // Exportieren des Katalogs in CSV-Dateien mit hierarchischer Struktur
        ExportToCsvWithHierarchy(catalog, outputDirectory);
        
    }

    /// <summary>
    /// Exportiert Produkte in CSV-Dateien mit einer hierarchischen Struktur basierend auf den Kategorien.
    /// </summary>
    /// <param name="catalog">Der zu exportierende Produktkatalog.</param>
    /// <param name="outputDirectory">Das Verzeichnis, in dem die CSV-Dateien gespeichert werden sollen.</param>
    public static void ExportToCsvWithHierarchy(ProductCatalog catalog, string outputDirectory)
    {
        // Sicherstellen, dass das Ausgabeverzeichnis existiert
        if (!Directory.Exists(outputDirectory))
        {
            Console.WriteLine("Erstelle Ausgabeverzeichnis...");
            Directory.CreateDirectory(outputDirectory);
        }

        Console.WriteLine($"Ausgabeverzeichnis: {Path.GetFullPath(outputDirectory)}");
        Console.WriteLine($"Anzahl der Kategorien: {catalog.CatalogStructures.Count}");
        Console.WriteLine($"Anzahl der Produkte: {catalog.Products.Count}");

        // Verarbeiten der Wurzelkategorien (Berücksichtigung von ParentId == "0")
        var potentialRootCategories = catalog.CatalogStructures.Where(cs => string.IsNullOrEmpty(cs.ParentId) || cs.ParentId == "0");

        if (!potentialRootCategories.Any())
        {
            Console.WriteLine("Keine Wurzelkategorien im Katalog gefunden.");
            return;
        }

        Console.WriteLine("Mögliche Wurzelkategorien:");
        foreach (var root in potentialRootCategories)
        {
            Console.WriteLine($"GroupId: {root.GroupId}, GroupName: {root.GroupName}, ParentId: {root.ParentId}");
        }

        // Verarbeitung jeder Wurzelkategorie
        foreach (var rootCategory in potentialRootCategories)
        {
            Console.WriteLine($"Verarbeite Wurzelkategorie: {rootCategory.GroupName}, GroupId: {rootCategory.GroupId}");
            string rootFolderPath = Path.Combine(outputDirectory, SanitizeFileName(rootCategory.GroupName));
            Directory.CreateDirectory(rootFolderPath);

            // Exportieren der Produkte auf Wurzelebene
            string rootFilePath = Path.Combine(rootFolderPath, $"{SanitizeFileName(rootCategory.GroupName)}.csv");
            Console.WriteLine($"Exportiere Produkte für Wurzelkategorie: {rootCategory.GroupName} nach {rootFilePath}");
            ExportProductsByCategory(catalog, rootCategory.GroupId, rootFilePath);

            foreach (var nodeCategory in catalog.CatalogStructures.Where(cs => cs.ParentId == rootCategory.GroupId))
            {
                Console.WriteLine($"  Verarbeite Knoten: {nodeCategory.GroupName}, GroupId: {nodeCategory.GroupId}");
                string nodeFolderPath = Path.Combine(rootFolderPath, SanitizeFileName(nodeCategory.GroupName));
                Directory.CreateDirectory(nodeFolderPath);

                // Exportieren der Produkte auf Knotenelementebene
                string nodeFilePath = Path.Combine(nodeFolderPath, $"{SanitizeFileName(nodeCategory.GroupName)}.csv");
                Console.WriteLine($"  Exportiere Produkte für Knoten: {nodeCategory.GroupName} nach {nodeFilePath}");
                ExportProductsByCategory(catalog, nodeCategory.GroupId, nodeFilePath);

                foreach (var leafCategory in catalog.CatalogStructures.Where(cs => cs.ParentId == nodeCategory.GroupId))
                {
                    Console.WriteLine($"    Verarbeite Blatt: {leafCategory.GroupName}, GroupId: {leafCategory.GroupId}");
                    string leafFilePath = Path.Combine(nodeFolderPath, $"{SanitizeFileName(leafCategory.GroupName)}.csv");

                    // Exportieren der Produkte auf Blattebene
                    Console.WriteLine($"    Exportiere Produkte für Blatt: {leafCategory.GroupName} nach {leafFilePath}");
                    ExportProductsByCategory(catalog, leafCategory.GroupId, leafFilePath);
                }
            }
        }

        Console.WriteLine("ExportToCsvWithHierarchy abgeschlossen.");
    }

    /// <summary>
    /// Exportiert Produkte einer bestimmten Kategorie in eine CSV-Datei.
    /// </summary>
    /// <param name="catalog">Der zu exportierende Produktkatalog.</param>
    /// <param name="categoryId">Die ID der Kategorie, deren Produkte exportiert werden sollen.</param>
    /// <param name="filePath">Der Pfad zur CSV-Datei, in die die Produkte exportiert werden sollen.</param>
    private static void ExportProductsByCategory(ProductCatalog catalog, string categoryId, string filePath)
    {
        Console.WriteLine($"Lade Produkte für Kategorie ID: {categoryId}");
        var products = catalog.Products
            .Where(p => p.ProductCatalogGroupMappings.Any(m => m.CatalogGroupId == categoryId))
            .ToList();

        if (!products.Any())
        {
            Console.WriteLine($"Keine Produkte gefunden für Kategorie ID: {categoryId}. Überspringe Datei: {filePath}");
            return;
        }

        Console.WriteLine($"Schreibe {products.Count} Produkte nach {filePath}");

        var csvContent = new StringBuilder();

        // Dynamisch Spalten für Features und Logistikdetails erstellen
        var allFeatureNames = products
            .SelectMany(p => p.FeatureSets ?? Enumerable.Empty<FeatureSet>())
            .SelectMany(fs => fs.Features)
            .Select(f => f.Name)
            .Distinct()
            .ToList();

        var logisticsColumns = new[] { "Length", "Width", "Depth", "Weight" };

        // Kopfzeile erstellen
        csvContent.Append("PRODUCT_NO;DESCRIPTION_SHORT;MANUFACTURER_NAME;NET_PRICE;GROSS_PRICE;KEYWORDS;");
        foreach (var featureName in allFeatureNames)
        {
            csvContent.Append($"{featureName};");
        }
        foreach (var column in logisticsColumns)
        {
            csvContent.Append($"{column};");
        }
        csvContent.AppendLine("ROOT_CATEGORY;NODE_CATEGORY;LEAF_CATEGORY");

        // Datenzeilen erstellen
        foreach (var product in products)
        {
            Console.WriteLine($"Verarbeite Produkt: {product.No}");
            var categories = GetCategoryHierarchy(catalog, product);

            var netPrice = product.Prices?.FirstOrDefault(p => p.PriceType == PriceTypes.NetList)?.Amount.ToString("F2") ?? "N/A";
            var grossPrice = product.Prices?.FirstOrDefault(p => p.PriceType == PriceTypes.NRP)?.Amount.ToString("F2") ?? "N/A";
            var keywords = product.Keywords != null && product.Keywords.Any()
                ? string.Join(", ", product.Keywords)
                : "N/A";

            csvContent.Append($"{product.No};{product.DescriptionShort};{product.ManufacturerName};{netPrice};{grossPrice};{keywords};");

            var productFeatures = product.FeatureSets?.SelectMany(fs => fs.Features).ToDictionary(f => f.Name, f => string.Join(", ", f.Values)) ?? new Dictionary<string, string>();
            foreach (var featureName in allFeatureNames)
            {
                csvContent.Append(productFeatures.ContainsKey(featureName) ? productFeatures[featureName] : "N/A");
                csvContent.Append(";");
            }

            var logistics = product.LogisticsDetails;
            csvContent.Append($"{logistics?.Length?.ToString() ?? "N/A"};");
            csvContent.Append($"{logistics?.Width?.ToString() ?? "N/A"};");
            csvContent.Append($"{logistics?.Depth?.ToString() ?? "N/A"};");
            csvContent.Append($"{logistics?.Weight?.ToString() ?? "N/A"};");

            csvContent.AppendLine($"{categories.Root};{categories.Node};{categories.Leaf}");
        }

        try
        {
            Console.WriteLine($"Schreibe CSV-Datei: {filePath}");
            File.WriteAllText(filePath, csvContent.ToString());
            Console.WriteLine($"Datei erfolgreich geschrieben: {filePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler beim Schreiben der Datei {filePath}: {ex.Message}");
        }
    }

    /// <summary>
    /// Ruft die Kategorienhierarchie (Root, Node, Leaf) für ein Produkt ab.
    /// </summary>
    /// <param name="catalog">Der Produktkatalog.</param>
    /// <param name="product">Das Produkt, für das die Hierarchie ermittelt werden soll.</param>
    /// <returns>Ein Tupel mit Root-, Node- und Leaf-Kategorie.</returns>
    private static (string Root, string Node, string Leaf) GetCategoryHierarchy(ProductCatalog catalog, Product product)
    {
        string root = "N/A", node = "N/A", leaf = "N/A";

        foreach (var mapping in product.ProductCatalogGroupMappings)
        {
            var leafStructure = catalog.CatalogStructures.FirstOrDefault(cs => cs.GroupId == mapping.CatalogGroupId);
            if (leafStructure != null)
            {
                leaf = leafStructure.GroupName;

                var nodeStructure = catalog.CatalogStructures.FirstOrDefault(cs => cs.GroupId == leafStructure.ParentId);
                if (nodeStructure != null)
                {
                    node = nodeStructure.GroupName;

                    var rootStructure = catalog.CatalogStructures.FirstOrDefault(cs => cs.GroupId == nodeStructure.ParentId);
                    if (rootStructure != null)
                    {
                        root = rootStructure.GroupName;
                    }
                }
            }
        }

        return (root, node, leaf);
    }

    /// <summary>
    /// Bereinigt Dateinamen von ungültigen Zeichen.
    /// </summary>
    /// <param name="fileName">Der ursprüngliche Dateiname.</param>
    /// <returns>Ein bereinigter Dateiname.</returns>
    private static string SanitizeFileName(string fileName)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            fileName = fileName.Replace(c, '_');
        }
        return fileName;
    }
 

}
