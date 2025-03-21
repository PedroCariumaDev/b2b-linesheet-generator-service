// excel-generator.js - Excel generation logic using template
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const axios = require('axios');

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
 * Optimize image URL by adding "_150x" before the extension
 * @param {string} imageUrl - Original image URL
 * @returns {string} - Optimized image URL
 */
function getOptimizedImageUrl(imageUrl) {
  if (!imageUrl || imageUrl.includes('/api/placeholder/')) {
    return imageUrl;
  }
  
  try {
    // Find the position of the last dot (extension)
    const lastDotIndex = imageUrl.lastIndexOf('.');
    if (lastDotIndex === -1) {
      return imageUrl; // No extension found
    }
    
    // Handle query parameters
    const queryIndex = imageUrl.indexOf('?', lastDotIndex);
    
    if (queryIndex === -1) {
      // No query parameters
      const base = imageUrl.substring(0, lastDotIndex);
      const ext = imageUrl.substring(lastDotIndex);
      return `${base}_150x${ext}`;
    } else {
      // Has query parameters
      const base = imageUrl.substring(0, lastDotIndex);
      const ext = imageUrl.substring(lastDotIndex, queryIndex);
      const query = imageUrl.substring(queryIndex);
      return `${base}_150x${ext}${query}`;
    }
  } catch (error) {
    console.error('Error optimizing image URL:', error);
    return imageUrl; // Return original on error
  }
}

/**
 * Fetch image as buffer from URL
 * @param {string} imageUrl - URL of the image to fetch
 * @returns {Promise<Buffer|null>} - Image buffer or null if failed
 */
async function fetchImageBuffer(imageUrl) {
  try {
    // Skip for placeholder images
    if (!imageUrl || imageUrl.includes('/api/placeholder/')) {
      console.log('Skipping placeholder image');
      return null;
    }

    // Optimize the image URL
    const optimizedUrl = getOptimizedImageUrl(imageUrl);
    console.log(`Original URL: ${imageUrl}`);
    console.log(`Optimized URL: ${optimizedUrl}`);
    
    console.log(`Fetching image from URL: ${optimizedUrl}`);
    
    const response = await axios.get(optimizedUrl, { 
      responseType: 'arraybuffer',
      timeout: 10000, // 10 second timeout
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36'
      }
    });
    
    console.log(`Image fetched successfully, size: ${response.data.length} bytes`);
    return Buffer.from(response.data, 'binary');
  } catch (error) {
    console.error(`Error fetching image:`, error.message);
    
    // Try the original URL as fallback
    if (!imageUrl.includes('_150x')) {
      try {
        console.log(`Trying fallback to original URL: ${imageUrl}`);
        const response = await axios.get(imageUrl, { 
          responseType: 'arraybuffer',
          timeout: 10000,
          headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36'
          }
        });
        console.log(`Fallback image fetched successfully, size: ${response.data.length} bytes`);
        return Buffer.from(response.data, 'binary');
      } catch (fallbackError) {
        console.error(`Error fetching original image:`, fallbackError.message);
      }
    }
    
    return null;
  }
}

/**
 * Generate Excel file(s) based on input data using template
 */
async function generateExcel(data) {
  const { company, catalogs, outputType } = data;
  console.log(`Starting Excel generation for ${company.name} with ${catalogs.length} catalogs`);
  
  // Path to template file
  const templatePath = path.join(__dirname, 'templates', 'B2B_Linesheet_BASE.xlsx');
  
  // Check if outputType is 'separate' and there are multiple catalogs
  if (outputType === 'separate' && catalogs.length > 1) {
    return await generateSeparateExcelFiles(company, catalogs, templatePath);
  }
  
  // Create a workbook from template
  const workbook = new ExcelJS.Workbook();
  
  try {
    // Try to read the template file
    await workbook.xlsx.readFile(templatePath);
    console.log('Successfully loaded template file');
  } catch (error) {
    console.error('Error loading template file:', error);
    console.log('Falling back to generating workbook from scratch');
    
    // Add metadata
    workbook.creator = 'Linesheet Generator';
    workbook.lastModifiedBy = company.name;
    workbook.created = new Date();
    workbook.modified = new Date();
    
    // Create a default template sheet
    workbook.addWorksheet('Winter 25');
    workbook.addWorksheet('Order Summary');
  }
  
  // Get template sheet (Winter 25) and Order Summary
  const templateSheet = workbook.getWorksheet('Winter 25');
  const summarySheet = workbook.getWorksheet('Order Summary');
  
  if (!templateSheet) {
    console.error('Template sheet "Winter 25" not found in the template file');
    return null;
  }
  
  if (!summarySheet) {
    console.error('Summary sheet "Order Summary" not found in the template file');
    return null;
  }
  
  // Fix the shared formula issue - convert all shared formulas to normal formulas
  // This function processes all sheets and converts shared formulas to normal ones
  function convertSharedFormulasToNormal(wb) {
    wb.eachSheet(sheet => {
      sheet.eachRow({ includeEmpty: false }, row => {
        row.eachCell({ includeEmpty: false }, cell => {
          if (cell.type === ExcelJS.ValueType.Formula) {
            // If it's a formula, get the formula text and set it again directly
            // This breaks the shared formula link
            const formulaText = cell.formula;
            if (formulaText) {
              // Store the value temporarily
              const oldValue = cell.value;
              
              // Reset the formula (but disconnect it from any shared formula)
              cell.value = { formula: formulaText };
              
              // If the formula produced a value, try to restore it
              if (oldValue && oldValue.result) {
                cell.result = oldValue.result;
              }
            }
          }
        });
      });
    });
  }
  
  // Convert any shared formulas in the template to normal formulas
  convertSharedFormulasToNormal(workbook);
  
  // Process each catalog - create a sheet for each using the template
  for (const catalog of catalogs) {
    // Create a new sheet based on the template
    const catalogSheetName = catalog.name;
    const catalogSheet = workbook.addWorksheet(catalogSheetName);
    
    console.log(`Creating sheet for catalog: ${catalogSheetName}`);
    
    // Copy content and formatting from template sheet but handle formulas carefully
    templateSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const newRow = catalogSheet.getRow(rowNumber);
      
      // Set row height to match template
      newRow.height = row.height;
      
      // Copy values and formatting from each cell
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const newCell = newRow.getCell(colNumber);
        
        // Handle formulas carefully
        if (cell.type === ExcelJS.ValueType.Formula) {
          // Copy the formula directly instead of copying the cell value object
          newCell.value = { formula: cell.formula };
          console.log(`Copied formula from cell ${cell.address} to ${newCell.address}: ${cell.formula}`);
        } else {
          // For non-formula cells, copy the value directly
          newCell.value = cell.value;
        }
        
        // Copy style as a separate operation
        if (cell.style) {
          try {
            newCell.style = JSON.parse(JSON.stringify(cell.style));
          } catch (styleError) {
            console.warn(`Could not copy style for cell ${cell.address}:`, styleError.message);
            // Apply minimal styling if JSON copy fails
            if (cell.style.font) newCell.font = cell.style.font;
            if (cell.style.alignment) newCell.alignment = cell.style.alignment;
            if (cell.style.border) newCell.border = cell.style.border;
            if (cell.style.fill) newCell.fill = cell.style.fill;
            if (cell.style.numFmt) newCell.numFmt = cell.style.numFmt;
          }
        }
      });
    });
    
    // Copy column properties (width, etc.)
    templateSheet.columns.forEach((col, index) => {
      if (col.width) {
        catalogSheet.getColumn(index + 1).width = col.width;
      }
    });
    
    // Set company and catalog info
    const companyCell = catalogSheet.getCell('B2'); // Retailer value
    const linesheetCell = catalogSheet.getCell('B3'); // Linesheet value
    
    companyCell.value = company.name;
    linesheetCell.value = catalog.name;
    
    // Add a simple test formula that doesn't depend on other cells
    try {
      const testFormulaCell = catalogSheet.getCell('D5');
      testFormulaCell.value = 10; // Set a plain number first
      testFormulaCell.value = { formula: '10+20' }; // Then set an absolute formula
      console.log(`Added test formula in cell D5: =10+20`);
    } catch (formulaError) {
      console.error('Error setting test formula:', formulaError);
    }
    
    // Clear any template/sample products
    // Hardcode the product start row based on template structure
    // This is more reliable than trying to detect it
    const productStartRow = 7; // First row of product data in template
    
    // Find all product rows in the template to clear them
    // We'll scan a significant number of rows to ensure we get all template data
    const rowsToScan = 25; // Scan more rows to ensure all template products are found
    
    console.log(`Scanning rows ${productStartRow} through ${productStartRow + rowsToScan - 1} for template products`);
    
    // First clear all potential product rows in the given range
    for (let i = 0; i < rowsToScan; i++) {
      const rowIndex = productStartRow + i;
      const row = catalogSheet.getRow(rowIndex);
      
      // Clear each cell in the row (up to column 30 to ensure all fields are cleared)
      for (let col = 1; col <= 30; col++) {
        const cell = row.getCell(col);
        // Clear the value but keep formatting
        cell.value = null;
      }
    }
    
    console.log(`Cleared ${rowsToScan} potential product rows starting at row ${productStartRow}`);

    // Add products from the catalog to the sheet
    if (catalog.products && catalog.products.length > 0) {
      console.log(`Adding ${catalog.products.length} products to ${catalogSheetName}`);
      
      // Set an appropriate row height for product rows with images
      const productRowHeight = 100; // Height in points, adjust as needed
      
      // Insert product data starting at productStartRow
      for (let i = 0; i < catalog.products.length; i++) {
        const product = catalog.products[i];
        const rowIndex = productStartRow + i;
        const row = catalogSheet.getRow(rowIndex);
        
        console.log(`Processing product ${i+1}/${catalog.products.length}: ${product.name || 'Unnamed'} (Row ${rowIndex})`);
        
        // Set row height to accommodate images
        row.height = productRowHeight;
        
        // Try to add the product image in column 0 (A)
        if (product.image && !product.image.includes('/api/placeholder/')) {
          try {
            console.log(`Attempting to fetch and embed image: ${product.image}`);
            
            const imageBuffer = await fetchImageBuffer(product.image);
            
            if (imageBuffer) {
              // Determine image extension from URL or default to png
              let extension = 'png';
              if (product.image.toLowerCase().endsWith('.jpg') || product.image.toLowerCase().endsWith('.jpeg')) {
                extension = 'jpeg';
              }
              
              const imageId = workbook.addImage({
                buffer: imageBuffer,
                extension: extension,
              });
              
              // Add the image to the cell
              catalogSheet.addImage(imageId, {
                tl: { col: 0, row: rowIndex - 1 }, // Top-left corner (0-indexed)
                br: { col: 1, row: rowIndex }, // Bottom-right corner
                editAs: 'oneCell', // Keeps the aspect ratio when row height changes
              });
              
              console.log(`Added image for product: ${product.name} with extension ${extension}`);
            } else {
              console.log(`Could not add image for product: ${product.name} (null buffer)`);
            }
          } catch (imageError) {
            console.error(`Error adding image for product ${product.name}:`, imageError.message);
          }
        } else {
          console.log(`Skipping image (invalid URL or placeholder): ${product.image || 'none'}`);
        }
        
        // Set product data in cells
        console.log('Setting product data cells');
        row.getCell(2).value = product.name || ''; // Product Name (column C)
        row.getCell(3).value = product.styleNumber || ''; // Style Number (column B)
        row.getCell(4).value = product.color || ''; // Color (column D)
        row.getCell(5).value = product.colorCode || ''; // Color Code (column E)
        row.getCell(6).value = product.season || ''; // Season (column F)
        row.getCell(7).value = product.evergreen || 'No'; // Evergreen (column G)
        row.getCell(8).value = product.countryOfOrigin || ''; // Country of Origin (column H)
        row.getCell(9).value = product.fabrication || ''; // Fabrication (column I)
        row.getCell(10).value = product.materialComposition || ''; // Material Composition (column J)
        row.getCell(11).value = product.category || ''; // Category (column K)
        row.getCell(12).value = product.subcategory || ''; // Subcategory (column L)
        row.getCell(13).value = product.sizeBreak || ''; // Size Break (column M)
        
        // Pricing - Format as currency
        const wholesalePriceCell = row.getCell(34); // Wholesale Price (column N)
        wholesalePriceCell.value = product.wholesalePrice || 0;
        wholesalePriceCell.numFmt = '$#,##0.00';
        
        const retailPriceCell = row.getCell(35); // Retail Price (column O)
        retailPriceCell.value = product.suggRetailPrice || 0;
        retailPriceCell.numFmt = '$#,##0.00';
        
        // Set values for sizes instead of adding formulas to avoid shared formula issues
        
        // Add size columns dynamically based on size break
        const sizeBreakNum = parseInt(product.sizeBreak, 10) || 1;
        const sizes = SIZE_BREAKS[sizeBreakNum] || [];
        
        // Add sizes in columns 17+ (starting from column Q)
        sizes.forEach((size, sizeIndex) => {
          const sizeCell = row.getCell(17 + sizeIndex);
          sizeCell.value = ''; // Default to empty, to be filled by customer for orders
          
          // Preserve column width and any formatting
          if (sizeCell.style) {
            sizeCell.style.alignment = { horizontal: 'center' };
          }
        });
        
        console.log(`Completed processing product: ${product.name || 'Unnamed'}`);
      }
    } else {
      console.log(`No products to add for catalog: ${catalogSheetName}`);
    }
  }
  
  // Update Order Summary sheet
  // TODO: Implement summary sheet logic
  
  // Remove the template sheet when done
  workbook.removeWorksheet(templateSheet.id);
  
  // Use the current workbook for the rest of the process
  const finalWorkbook = workbook;
  
  // Convert shared formulas one more time before writing
  convertSharedFormulasToNormal(finalWorkbook);
  
  // Create filename
  let filename = `${company.name.replace(/\s+/g, '_')}_Linesheet.xlsx`;
  
  // Write to buffer
  console.log('Generating Excel buffer...');
  try {
    const buffer = await finalWorkbook.xlsx.writeBuffer();
    console.log(`Excel generation complete: ${filename}`);
    return { buffer, filename };
  } catch (writeError) {
    console.error('Error writing Excel buffer:', writeError);
    
    // Try a more desperate approach - remove all formulas and just keep values
    console.log('Attempting to remove problematic formulas...');
    finalWorkbook.eachSheet(sheet => {
      sheet.eachRow({ includeEmpty: false }, row => {
        row.eachCell({ includeEmpty: false }, cell => {
          if (cell.type === ExcelJS.ValueType.Formula) {
            // Replace formula with its value if possible
            if (cell.value && cell.value.result !== undefined) {
              cell.value = cell.value.result;
            } else {
              // Otherwise just clear the formula
              cell.value = null;
            }
          }
        });
      });
    });

    // Try writing again
    console.log('Retrying Excel generation after formula cleanup...');
    const buffer = await finalWorkbook.xlsx.writeBuffer();
    console.log(`Excel generation complete after cleanup: ${filename}`);
    return { buffer, filename };
  }
}

/**
 * Generate separate Excel files for each catalog
 */
async function generateSeparateExcelFiles(company, catalogs, templatePath) {
  console.log(`Generating separate Excel files for ${catalogs.length} catalogs`);
  
  // Create a result object with array of files
  const result = {
    files: [],
    outputType: 'separate'
  };
  
  // Process each catalog as a separate file
  for (const catalog of catalogs) {
    // Create a new workbook from template for each catalog
    const workbook = new ExcelJS.Workbook();
    
    try {
      // Read the template file
      await workbook.xlsx.readFile(templatePath);
      
      // Fix shared formulas immediately after loading
      workbook.eachSheet(sheet => {
        sheet.eachRow({ includeEmpty: false }, row => {
          row.eachCell({ includeEmpty: false }, cell => {
            if (cell.type === ExcelJS.ValueType.Formula) {
              const formulaText = cell.formula;
              if (formulaText) {
                cell.value = { formula: formulaText };
              }
            }
          });
        });
      });
    } catch (error) {
      console.error(`Error loading template for catalog ${catalog.name}:`, error);
      continue; // Skip this catalog and move to the next
    }
    
    // Get template sheet (Winter 25) and Order Summary
    const templateSheet = workbook.getWorksheet('Winter 25');
    const summarySheet = workbook.getWorksheet('Order Summary');
    
    if (!templateSheet || !summarySheet) {
      console.error(`Required sheets not found for catalog ${catalog.name}`);
      continue; // Skip this catalog
    }
    
    // Rename template sheet to catalog name
    templateSheet.name = catalog.name;
    
    // Set company and catalog info
    const companyCell = templateSheet.getCell('B2'); // Retailer value
    const linesheetCell = templateSheet.getCell('B3'); // Linesheet value
    
    companyCell.value = company.name;
    linesheetCell.value = catalog.name;
    
    // Define product start row and clear all potential template rows
    const productStartRow = 8; // First row of product data
    const rowsToScan = 25; // Clear a good number of rows to ensure all template products are gone
    
    console.log(`Clearing ${rowsToScan} potential template product rows in catalog ${catalog.name}`);
    
    // Clear all potential product rows
    for (let i = 0; i < rowsToScan; i++) {
      const rowIndex = productStartRow + i;
      const row = templateSheet.getRow(rowIndex);
      
      // Clear all cells in this row (columns 1-30 to be safe)
      for (let col = 1; col <= 30; col++) {
        const cell = row.getCell(col);
        // Preserve formatting but clear value
        cell.value = null;
      }
    }
    
    // Add products from the catalog
    if (catalog.products && catalog.products.length > 0) {
      console.log(`Adding ${catalog.products.length} products to catalog ${catalog.name}`);
      
      // Set an appropriate row height for product rows with images
      const productRowHeight = 100; // Height in points, adjust as needed
      
      // Process each product
      for (let i = 0; i < catalog.products.length; i++) {
        const product = catalog.products[i];
        const rowIndex = productStartRow + i;
        const row = templateSheet.getRow(rowIndex);
        
        console.log(`Separate file - Processing product ${i+1}/${catalog.products.length}: ${product.name || 'Unnamed'}`);
        
        // Set row height to accommodate images
        row.height = productRowHeight;
        
        // Try to add the product image in column A
        if (product.image && !product.image.includes('/api/placeholder/')) {
          try {
            console.log(`Attempting to fetch and embed image: ${product.image}`);
            
            const imageBuffer = await fetchImageBuffer(product.image);
            
            if (imageBuffer) {
              // Determine image extension from URL or default to png
              let extension = 'png';
              if (product.image.toLowerCase().endsWith('.jpg') || product.image.toLowerCase().endsWith('.jpeg')) {
                extension = 'jpeg';
              }
              
              const imageId = workbook.addImage({
                buffer: imageBuffer,
                extension: extension,
              });
              
              // Add the image to the cell
              templateSheet.addImage(imageId, {
                tl: { col: 0, row: rowIndex - 1 }, // Top-left corner (0-indexed)
                br: { col: 1, row: rowIndex }, // Bottom-right corner
                editAs: 'oneCell', // Keeps the aspect ratio when row height changes
              });
              
              console.log(`Added image for product: ${product.name} with extension ${extension}`);
            } else {
              console.log(`Could not add image for product: ${product.name}`);
            }
          } catch (imageError) {
            console.error(`Error adding image for product ${product.name}:`, imageError.message);
          }
        } else {
          console.log(`Skipping image (invalid URL or placeholder): ${product.image || 'none'}`);
        }
        
        // Set product data in cells
        row.getCell(2).value = product.styleNumber || ''; // Style Number
        row.getCell(3).value = product.name || ''; // Product Name
        row.getCell(4).value = product.color || ''; // Color
        row.getCell(5).value = product.colorCode || ''; // Color Code
        row.getCell(6).value = product.season || ''; // Season
        row.getCell(7).value = product.evergreen || 'No'; // Evergreen
        row.getCell(8).value = product.countryOfOrigin || ''; // Country of Origin
        row.getCell(9).value = product.fabrication || ''; // Fabrication
        row.getCell(10).value = product.materialComposition || ''; // Material Composition
        row.getCell(11).value = product.category || ''; // Category
        row.getCell(12).value = product.subcategory || ''; // Subcategory
        row.getCell(13).value = product.sizeBreak || '1'; // Size Break
        
        // Pricing
        const wholesalePriceCell = row.getCell(14);
        wholesalePriceCell.value = product.wholesalePrice || 0;
        wholesalePriceCell.numFmt = '$#,##0.00';
        
        const retailPriceCell = row.getCell(15);
        retailPriceCell.value = product.suggRetailPrice || 0;
        retailPriceCell.numFmt = '$#,##0.00';
        
        // Avoid adding dependent formulas to prevent shared formula issues
        
        // Add size columns based on size break
        const sizeBreakNum = parseInt(product.sizeBreak, 10) || 1;
        const sizes = SIZE_BREAKS[sizeBreakNum] || [];
        
        sizes.forEach((size, sizeIndex) => {
          const sizeCell = row.getCell(17 + sizeIndex);
          sizeCell.value = ''; // Empty by default
        });
      }
    }
    
    // Create a sanitized filename for this catalog
    const filename = `${company.name.replace(/\s+/g, '_')}_${catalog.name.replace(/\s+/g, '_')}.xlsx`;
    
    // Write to buffer
    console.log(`Generating file buffer for: ${filename}`);
    try {
      const buffer = await workbook.xlsx.writeBuffer();
      console.log(`Created file: ${filename}`);
      
      // Add to result files array
      result.files.push({
        buffer,
        filename,
        catalogId: catalog.id
      });
    } catch (writeError) {
      console.error(`Error writing Excel buffer for catalog ${catalog.name}:`, writeError);
      
      // Try a more desperate approach - remove all formulas and just keep values
      console.log('Attempting to remove problematic formulas...');
      workbook.eachSheet(sheet => {
        sheet.eachRow({ includeEmpty: false }, row => {
          row.eachCell({ includeEmpty: false }, cell => {
            if (cell.type === ExcelJS.ValueType.Formula) {
              // Replace formula with its value if possible
              if (cell.value && cell.value.result !== undefined) {
                cell.value = cell.value.result;
              } else {
                // Otherwise just clear the formula
                cell.value = null;
              }
            }
          });
        });
      });
      
      // Try writing again
      console.log(`Retrying Excel generation for catalog ${catalog.name} after formula cleanup...`);
      try {
        const buffer = await workbook.xlsx.writeBuffer();
        console.log(`Created file after cleanup: ${filename}`);
        
        // Add to result files array
        result.files.push({
          buffer,
          filename,
          catalogId: catalog.id
        });
      } catch (retryError) {
        console.error(`Failed to generate Excel for catalog ${catalog.name} even after cleanup:`, retryError);
      }
    }
  }
  
  return result;
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