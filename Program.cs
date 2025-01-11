using BMECat.net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

class Program
{
    static void Main(string[] args)
    {
        string filePath = "picard_bmecat_de.xml";
        string outputDirectory = "OutputCsv";

        // Load the catalog
        ProductCatalog catalog = ProductCatalog.Load(filePath);

        if (catalog == null || catalog.Products == null || catalog.Products.Count == 0)
        {
            Console.WriteLine("The catalog is empty or not loaded.");
            return;
        }

        // Export products to CSVs with hierarchy
        ExportToCsvWithHierarchy(catalog, outputDirectory);
        

        Console.WriteLine($"Export completed. Check the directory: {Path.GetFullPath(outputDirectory)}");

        // Add debug logging
foreach (var root in catalog.CatalogStructures.Where(cs => string.IsNullOrEmpty(cs.ParentId)))
{
    Console.WriteLine($"Processing Root: {root.GroupName}, Group ID: {root.GroupId}");

    foreach (var node in catalog.CatalogStructures.Where(cs => cs.ParentId == root.GroupId))
    {
        Console.WriteLine($"  Processing Node: {node.GroupName}, Group ID: {node.GroupId}");

        foreach (var leaf in catalog.CatalogStructures.Where(cs => cs.ParentId == node.GroupId))
        {
            Console.WriteLine($"    Processing Leaf: {leaf.GroupName}, Group ID: {leaf.GroupId}");
        }
    }
}

    }

    public static void ExportToCsvWithHierarchy(ProductCatalog catalog, string outputDirectory)
{
    // Ensure the output directory exists
    if (!Directory.Exists(outputDirectory))
    {
        Directory.CreateDirectory(outputDirectory);
    }

    Console.WriteLine($"Output Directory: {Path.GetFullPath(outputDirectory)}");

    // Find all root categories
    foreach (var root in catalog.CatalogStructures.Where(cs => string.IsNullOrEmpty(cs.ParentId)))
    {
        string rootFolderPath = Path.Combine(outputDirectory, SanitizeFileName(root.GroupName));
        Directory.CreateDirectory(rootFolderPath);

        Console.WriteLine($"Processing Root: {root.GroupName} (ID: {root.GroupId})");

        // Export products at the root level
        ExportProductsByCategory(catalog, root.GroupId, Path.Combine(rootFolderPath, $"{SanitizeFileName(root.GroupName)}.csv"));

        // Find all nodes under this root
        foreach (var node in catalog.CatalogStructures.Where(cs => cs.ParentId == root.GroupId))
        {
            string nodeFolderPath = Path.Combine(rootFolderPath, SanitizeFileName(node.GroupName));
            Directory.CreateDirectory(nodeFolderPath);

            Console.WriteLine($"  Processing Node: {node.GroupName} (ID: {node.GroupId})");

            // Export products at the node level
            ExportProductsByCategory(catalog, node.GroupId, Path.Combine(nodeFolderPath, $"{SanitizeFileName(node.GroupName)}.csv"));

            // Find all leaves under this node
            foreach (var leaf in catalog.CatalogStructures.Where(cs => cs.ParentId == node.GroupId))
            {
                Console.WriteLine($"    Processing Leaf: {leaf.GroupName} (ID: {leaf.GroupId})");

                // Export products at the leaf level
                ExportProductsByCategory(catalog, leaf.GroupId, Path.Combine(nodeFolderPath, $"{SanitizeFileName(leaf.GroupName)}.csv"));
            }
        }
    }
}

private static void ExportProductsByCategory(ProductCatalog catalog, string categoryId, string filePath)
{
    // Find products mapped to this category ID
    var products = catalog.Products
        .Where(p => p.ProductCatalogGroupMappings.Any(m => m.CatalogGroupId == categoryId))
        .ToList();

    // Log product count
    if (products.Count == 0)
    {
        Console.WriteLine($"  No products found for category ID: {categoryId}. Skipping file: {filePath}");
        return;
    }

    Console.WriteLine($"  Writing {products.Count} products to {filePath}");

    // Build CSV content
    StringBuilder csvContent = new StringBuilder();
    csvContent.AppendLine("PRODUCT_NO;DESCRIPTION_SHORT;MANUFACTURER_NAME;ROOT_CATEGORY;NODE_CATEGORY;LEAF_CATEGORY");

    foreach (var product in products)
    {
        var categories = GetCategoryHierarchy(catalog, product);
        csvContent.AppendLine($"{product.No};{product.DescriptionShort};{product.ManufacturerName};{categories.Root};{categories.Node};{categories.Leaf}");
    }

    try
    {
        // Write CSV to file
        File.WriteAllText(filePath, csvContent.ToString(), Encoding.UTF8);
        Console.WriteLine($"  File successfully written: {filePath}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"  Error writing file {filePath}: {ex.Message}");
    }
}

private static string SanitizeFileName(string name)
{
    return string.Concat(name.Split(Path.GetInvalidFileNameChars()));
}

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

}
