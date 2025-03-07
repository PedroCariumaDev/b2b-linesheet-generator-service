// excel-generator.js - Excel generation logic using template
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

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
 * Generate Excel file(s) based on input data using template
 */
async function generateExcel(data) {
  const { company, catalogs, outputType } = data;
  
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
  
  // Process each catalog - create a sheet for each using the template
  for (const catalog of catalogs) {
    // Create a new sheet based on the template
    const catalogSheetName = catalog.name;
    const catalogSheet = workbook.addWorksheet(catalogSheetName);
    
    // Copy content and formatting from template sheet
    templateSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const newRow = catalogSheet.getRow(rowNumber);
      
      // Copy values and formatting from each cell
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const newCell = newRow.getCell(colNumber);
        
        // Copy value
        newCell.value = cell.value;
        
        // Copy style
        newCell.style = JSON.parse(JSON.stringify(cell.style));
        
        // Handle formulas specially
        if (cell.formula) {
          newCell.formula = cell.formula;
        }
      });
      
      // Set row height to match template
      newRow.height = row.height;
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
    
    // We'll handle adding products later
  }
  
  // Keep the Order Summary sheet as is for now
  // We'll implement the product data update in the next iteration
  
  // Remove the template sheet when done
  workbook.removeWorksheet(templateSheet.id);
  
  // Explicitly set sheet order by creating a new workbook with sheets in the correct order
  const orderedWorkbook = new ExcelJS.Workbook();
  
  // First, add all catalog sheets in their original order
  for (const worksheet of workbook.worksheets) {
    if (worksheet.name !== 'Order Summary') {
      // Clone the worksheet to the new workbook
      const newSheet = orderedWorkbook.addWorksheet(worksheet.name, {
        properties: JSON.parse(JSON.stringify(worksheet.properties)),
        pageSetup: worksheet.pageSetup,
        views: worksheet.views,
        state: worksheet.state
      });
      
      // Copy column properties
      worksheet.columns.forEach((col, index) => {
        if (col.width) {
          newSheet.getColumn(index + 1).width = col.width;
        }
      });
      
      // Copy content and formatting
      worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        const newRow = newSheet.getRow(rowNumber);
        newRow.height = row.height;
        
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const newCell = newRow.getCell(colNumber);
          
          // Copy value
          if (cell.formula) {
            newCell.formula = cell.formula;
          } else {
            newCell.value = cell.value;
          }
          
          // Copy style
          if (cell.style) {
            newCell.style = JSON.parse(JSON.stringify(cell.style));
          }
        });
      });
    }
  }
  
  // Finally add the Order Summary sheet
  if (summarySheet) {
    const newSummarySheet = orderedWorkbook.addWorksheet('Order Summary', {
      properties: JSON.parse(JSON.stringify(summarySheet.properties)),
      pageSetup: summarySheet.pageSetup,
      views: summarySheet.views,
      state: summarySheet.state
    });
    
    // Copy column widths
    summarySheet.columns.forEach((col, index) => {
      if (col.width) {
        newSummarySheet.getColumn(index + 1).width = col.width;
      }
    });
    
    // Copy content and formatting
    summarySheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const newRow = newSummarySheet.getRow(rowNumber);
      newRow.height = row.height;
      
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const newCell = newRow.getCell(colNumber);
        
        // Copy value
        if (cell.formula) {
          newCell.formula = cell.formula;
        } else {
          newCell.value = cell.value;
        }
        
        // Copy style
        if (cell.style) {
          newCell.style = JSON.parse(JSON.stringify(cell.style));
        }
      });
    });
  }
  
  // Use the ordered workbook for the rest of the process
  const finalWorkbook = orderedWorkbook;
  
  // Create filename
  let filename = `${company.name.replace(/\s+/g, '_')}_Linesheet.xlsx`;
  
  // Write to buffer
  const buffer = await finalWorkbook.xlsx.writeBuffer();
  
  return { buffer, filename };
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
    
    // Create a sanitized filename for this catalog
    const filename = `${company.name.replace(/\s+/g, '_')}_${catalog.name.replace(/\s+/g, '_')}.xlsx`;
    
    // Add products to the sheet here (We'll implement this in the next iteration)
    
    // Write to buffer
    const buffer = await workbook.xlsx.writeBuffer();
    
    // Add to result files array
    result.files.push({
      buffer,
      filename,
      catalogId: catalog.id
    });
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