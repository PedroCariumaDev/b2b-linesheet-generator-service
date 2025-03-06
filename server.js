// server.js - Main server file
const express = require('express');
const cors = require('cors');
const dotenv = require('dotenv');
const generateExcel = require('./excel-generator');
const shopifyApi = require('./shopify-api');

// Load environment variables
dotenv.config();

const app = express();
const port = process.env.PORT || 3000;
const allowedOrigin = process.env.ALLOWED_ORIGIN || '*';

// Enable CORS and JSON parsing
app.use(cors({
  origin: allowedOrigin,
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));
app.use(express.json({ limit: '10mb' })); // Increase limit for larger data

// Test endpoint
app.get('/', (req, res) => {
  res.send('Linesheet Generator API is running');
});

/**
 * Fetch B2B catalogs for a customer
 * GET /api/catalogs/:customerId
 */
app.get('/api/catalogs/:customerId', async (req, res) => {
  try {
    const { customerId } = req.params;
    console.log(`Fetching catalogs for customer: ${customerId}`);
    
    const catalogs = await shopifyApi.fetchB2BCatalogs(customerId);
    res.json(catalogs);
  } catch (error) {
    console.error('Error fetching catalogs:', error);
    res.status(500).json({
      error: 'Failed to fetch catalogs',
      message: error.message
    });
  }
});

/**
 * Fetch products for a catalog
 * GET /api/catalogs/:catalogId/products
 */
app.get('/api/catalogs/:catalogId/products', async (req, res) => {
  try {
    const { catalogId } = req.params;
    console.log(`Fetching products for catalog: ${catalogId}`);
    
    const products = await shopifyApi.fetchCatalogProducts(catalogId);
    res.json(products);
  } catch (error) {
    console.error('Error fetching catalog products:', error);
    res.status(500).json({
      error: 'Failed to fetch catalog products',
      message: error.message
    });
  }
});

/**
 * Fetch complete data for multiple catalogs (catalogs + products)
 * GET /api/complete-data/:customerId
 */
app.get('/api/complete-data/:customerId', async (req, res) => {
  try {
    const { customerId } = req.params;
    console.log(`Fetching complete data for customer: ${customerId}`);
    
    // Get catalogs for the customer
    const catalogs = await shopifyApi.fetchB2BCatalogs(customerId);
    
    // Fetch products for each catalog
    for (const catalog of catalogs) {
      catalog.products = await shopifyApi.fetchCatalogProducts(catalog.id);
    }
    
    // Return the complete data
    res.json({
      customerId,
      catalogs
    });
  } catch (error) {
    console.error('Error fetching complete data:', error);
    res.status(500).json({
      error: 'Failed to fetch complete data',
      message: error.message
    });
  }
});

// Excel generation endpoint
app.post('/api/generate-linesheet', async (req, res) => {
  try {
    console.log('Received request to generate linesheet');
    
    const { company, catalogIds, catalogs, outputType } = req.body;
    
    // Validate input
    if (!company || !catalogs || !Array.isArray(catalogs) || catalogs.length === 0) {
      return res.status(400).json({ error: 'Invalid input data' });
    }
    
    console.log(`Generating ${outputType} linesheet for ${company.name}`);
    console.log(`Selected catalogs: ${catalogIds.join(', ')}`);
    
    // Generate Excel file(s)
    const result = await generateExcel(req.body);
    
    // Set response headers for file download
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${result.filename}"`);
    
    // Send the Excel file as the response
    res.send(result.buffer);
    
  } catch (error) {
    console.error('Error generating Excel:', error);
    res.status(500).json({ error: 'Failed to generate Excel file' });
  }
});

// Start the server
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
  console.log(`Allowing requests from: ${allowedOrigin}`);
});