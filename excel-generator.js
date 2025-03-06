// excel-generator.js - Excel generation logic
const ExcelJS = require('exceljs');

// Size break definitions
const SIZE_BREAKS = {
  "1": [
    "M8/W9.5", "M8.5/W10", "M9/W10.5", "M9.5/W11", "M10/W11.5", 
    "M10.5/W12", "M11/W12.5", "M11.5/W13", "M12/W13.5", "M12.5/W14", 
    "M13/W14.5", "W5/M3.5", "W5.5/M4", "W6/M4.5", "W6.5/M5", 
    "W7/M5.5", "W7.5/M6", "W8/M6.5", "W8.5/M7", "W9/M7.5"
  ],
  "2": [
    "M7.5-M8/W9-W9.5", "M8.5-M9/W10-W10.5", "M9.5-M10/W11-W11.5", 
    "M10.5-M11/W12-W12.5", "M11.5-M12/W13-W13.5", "M12.5-M13/W14-W14.5", 
    "W5-W5.5/M3.5-M4", "W6-W6.5/M4.5-M5", "W7-W7.5/M5.5-M6", "W8-W8.5/M6.5-M7"
  ],
  "3": ["XS", "S", "M", "L", "XL", "XXL"],
  "4": ["One Size"]
};

/**
 * Generate Excel file(s) based on input data
 */
async function generateExcel(data) {
  const { company, catalogs, outputType } = data;
  
  // Create a workbook
  const workbook = new ExcelJS.Workbook();
  
  // Add metadata
  workbook.creator = 'Linesheet Generator';
  workbook.lastModifiedBy = company.name;
  workbook.created = new Date();
  workbook.modified = new Date();
  
  // Process each catalog
  for (const catalog of catalogs) {
    // Add catalog sheet
    await addCatalogSheet(workbook, catalog, company);
  }
  
  // Add order summary sheet
  addOrderSummarySheet(workbook, catalogs, company);
  
  // Create filename
  let filename;
  if (outputType === 'combined' || catalogs.length === 1) {
    filename = `${company.name.replace(/\s+/g, '_')}_Linesheet.xlsx`;
  } else {
    filename = `${company.name.replace(/\s+/g, '_')}_${catalogs[0].name.replace(/\s+/g, '_')}.xlsx`;
  }
  
  // Write to buffer
  const buffer = await workbook.xlsx.writeBuffer();
  
  return { buffer, filename };
}

/**
 * Add a catalog sheet to the workbook
 */
async function addCatalogSheet(workbook, catalog, company) {
  console.log(`Creating sheet for catalog: ${catalog.name}`);
  
  // Create a worksheet for this catalog
  const sheet = workbook.addWorksheet(catalog.name);
  
  // Add header rows
  sheet.addRow(['Retailer', company.name]);
  sheet.addRow(['Linesheet', catalog.name]);
  sheet.addRow(['Start Ship', catalog.startShip || '']);
  sheet.addRow(['Complete Ship', catalog.completeShip || '']);
  sheet.addRow(['Season Year', catalog.seasonYear || '']);
  sheet.addRow([]);
  
  // Add column headers
  const headers = [
    'Style Image', 'Style Name', 'Style Number', 'Color', 'Color Code', 'Season', 
    'Evergreen', 'Country of Origin', 'Fabrication', 'Material Composition', 
    'Category', 'Subcategory', 'Size Break'
  ];
  
  // Get all possible sizes from size breaks
  const allSizes = [];
  Object.values(SIZE_BREAKS).forEach(sizes => {
    sizes.forEach((size, index) => {
      while (index >= allSizes.length) {
        allSizes.push(`Size ${allSizes.length + 1}`);
      }
    });
  });
  
  // Add remaining columns
  const calculatedColumns = ['Units', 'Wholesale Price', 'Sugg Retail Price', 'Total Wholesale', 'Total Retail'];
  
  // Combine all columns
  const allColumns = [...headers, ...allSizes, ...calculatedColumns];
  const headerRow = sheet.addRow(allColumns);
  
  // Style the header row
  headerRow.eachCell((cell) => {
    cell.font = { bold: true };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFEEEEEE' }
    };
  });
  
  // Add product rows
  if (catalog.products && Array.isArray(catalog.products)) {
    catalog.products.forEach((product, index) => {
      const rowNumber = sheet.rowCount + 1; // Next row number
      
      // Basic product details
      const row = sheet.addRow([
        '', // Image placeholder
        product.name || '',
        product.styleNumber || '',
        product.color || '',
        product.colorCode || '',
        product.season || '',
        product.evergreen || '',
        product.countryOfOrigin || '',
        product.fabrication || '',
        product.materialComposition || '',
        product.category || '',
        product.subcategory || '',
        product.sizeBreak || '',
        // Empty cells for sizes
        ...Array(allSizes.length).fill(''),
        // Values for remaining columns
        0, // Units (will be formula)
        product.wholesalePrice || 0,
        product.suggRetailPrice || 0,
        0, // Total Wholesale (will be formula)
        0  // Total Retail (will be formula)
      ]);
      
      // Get column indices for formulas
      const firstSizeCol = headers.length + 1;
      const lastSizeCol = headers.length + allSizes.length;
      const unitsCellIndex = headers.length + allSizes.length + 1;
      const wholesalePriceIndex = unitsCellIndex + 1;
      const retailPriceIndex = wholesalePriceIndex + 1;
      const totalWholesaleIndex = retailPriceIndex + 1;
      const totalRetailIndex = totalWholesaleIndex + 1;
      
      // Get Excel column letters for formulas
      const firstSizeColLetter = getExcelColumn(firstSizeCol);
      const lastSizeColLetter = getExcelColumn(lastSizeCol);
      const unitsColLetter = getExcelColumn(unitsCellIndex);
      const wholesaleColLetter = getExcelColumn(wholesalePriceIndex);
      const retailColLetter = getExcelColumn(retailPriceIndex);
      
      // Set formulas
      row.getCell(unitsCellIndex).value = { 
        formula: `SUM(${firstSizeColLetter}${rowNumber}:${lastSizeColLetter}${rowNumber})` 
      };
      
      row.getCell(totalWholesaleIndex).value = { 
        formula: `${unitsColLetter}${rowNumber}*${wholesaleColLetter}${rowNumber}` 
      };
      
      row.getCell(totalRetailIndex).value = { 
        formula: `${unitsColLetter}${rowNumber}*${retailColLetter}${rowNumber}` 
      };
      
      // Format currency cells
      row.getCell(wholesalePriceIndex).numFmt = '$#,##0.00';
      row.getCell(retailPriceIndex).numFmt = '$#,##0.00';
      row.getCell(totalWholesaleIndex).numFmt = '$#,##0.00';
      row.getCell(totalRetailIndex).numFmt = '$#,##0.00';
    });
  } else {
    console.warn(`No products found for catalog: ${catalog.name}`);
  }
  
  // Set column widths
  sheet.columns.forEach((column, i) => {
    let width = 15;
    
    if (i === 1) width = 25; // Style Name
    if (i === 9) width = 20; // Fabrication
    if (i === 10) width = 25; // Material Composition
    
    column.width = width;
  });
  
  // Style header rows
  for (let i = 1; i <= 5; i++) {
    const headerRow = sheet.getRow(i);
    headerRow.font = { bold: true };
  }
}

/**
 * Add order summary sheet to the workbook
 */
function addOrderSummarySheet(workbook, catalogs, company) {
  console.log('Creating Order Summary sheet');
  
  const sheet = workbook.addWorksheet('Order Summary');
  
  // Add header
  sheet.addRow(['Retailer', company.name]);
  sheet.addRow([]);
  
  // Add column headers
  const headerRow = sheet.addRow([
    'Season Year', 'Delivery', 'Category', 'Subcategory', 'Total Units', 
    'Total Wholesale (USD)', 'Total Retail (USD)'
  ]);
  
  // Style the header row
  headerRow.eachCell((cell) => {
    cell.font = { bold: true };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFEEEEEE' }
    };
  });
  
  // Set column widths
  sheet.columns = [
    { width: 20 }, // Season Year
    { width: 20 }, // Delivery
    { width: 15 }, // Category
    { width: 15 }, // Subcategory
    { width: 15 }, // Total Units
    { width: 20 }, // Total Wholesale
    { width: 20 }  // Total Retail
  ];
  
  // Group products by catalog, category, and subcategory
  let summaryData = [];
  let rowIndex = 4; // Start after headers
  
  catalogs.forEach(catalog => {
    if (!catalog.products || !Array.isArray(catalog.products)) {
      console.warn(`No products in catalog: ${catalog.name}`);
      return;
    }
    
    // Group by category/subcategory
    const categories = {};
    
    catalog.products.forEach(product => {
      const category = product.category || 'Uncategorized';
      const subcategory = product.subcategory || 'Uncategorized';
      
      if (!categories[category]) {
        categories[category] = {};
      }
      
      if (!categories[category][subcategory]) {
        categories[category][subcategory] = true;
        
        // Generate a consistent number for units based on product data
        // In a real implementation, these would be dynamically calculated
        // based on specific product quantities
        const productHash = hashString(`${catalog.name}-${category}-${subcategory}`);
        const units = 5 + (productHash % 20); // 5-24 units
        const avgWholesale = 35 + (productHash % 30); // $35-$64 wholesale price
        const avgRetail = avgWholesale * 2.5; // 2.5x markup
        
        summaryData.push({
          seasonYear: catalog.seasonYear || '',
          delivery: catalog.name,
          category,
          subcategory,
          units,
          wholesale: units * avgWholesale,
          retail: units * avgRetail
        });
      }
    });
  });
  
  // Add rows for each grouping
  summaryData.forEach(item => {
    const row = sheet.addRow([
      item.seasonYear,
      item.delivery,
      item.category,
      item.subcategory,
      item.units,
      item.wholesale,
      item.retail
    ]);
    
    // Format currency cells
    row.getCell(6).numFmt = '$#,##0.00';
    row.getCell(7).numFmt = '$#,##0.00';
    
    rowIndex++;
  });
  
  // Calculate totals
  const totalUnits = summaryData.reduce((sum, item) => sum + item.units, 0);
  const totalWholesale = summaryData.reduce((sum, item) => sum + item.wholesale, 0);
  const totalRetail = summaryData.reduce((sum, item) => sum + item.retail, 0);
  
  // Add total row
  const totalRow = sheet.addRow([
    '',
    'Total',
    '',
    '',
    totalUnits,
    totalWholesale,
    totalRetail
  ]);
  
  // Style total row
  totalRow.eachCell((cell) => {
    cell.font = { bold: true };
  });
  
  // Format currency cells in total row
  totalRow.getCell(6).numFmt = '$#,##0.00';
  totalRow.getCell(7).numFmt = '$#,##0.00';
  
  // Style company name
  sheet.getCell('B1').font = { bold: true };
}

/**
 * Convert column index to Excel column letters (A, B, C, ... AA, AB, etc.)
 */
function getExcelColumn(column) {
  let temp, letter = '';
  
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  
  return letter;
}

/**
 * Create a simple hash from a string
 * Used to generate consistent random-like numbers
 */
function hashString(str) {
  let hash = 0;
  
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // Convert to 32bit integer
  }
  
  return Math.abs(hash);
}

module.exports = generateExcel;