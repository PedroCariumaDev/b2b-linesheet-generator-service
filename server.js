// server.js - Main server file
const express = require('express');
const cors = require('cors');
const dotenv = require('dotenv');
const archiver = require('archiver'); // Add this for ZIP functionality
const generateExcel = require('./excel-generator');
const shopifyApi = require('./shopify-api');

// Load environment variables
dotenv.config();

const app = express();
const port = process.env.PORT || 3000;
const allowedOrigins = [
  'https://test-cariuma.myshopify.com', 
  'https://*.myshopify.com',  // Allow all Shopify stores
  process.env.ALLOWED_ORIGIN || '*'
];

// Enable CORS and JSON parsing
app.use(cors({
  origin: function(origin, callback) {
    // Allow requests with no origin (like mobile apps or curl requests)
    if (!origin) return callback(null, true);
    
    // Check if the origin is allowed
    const isAllowed = allowedOrigins.some(allowedOrigin => {
      // Handle wildcard domains
      if (allowedOrigin.includes('*')) {
        const pattern = new RegExp(allowedOrigin.replace('*', '.*'));
        return pattern.test(origin);
      }
      return allowedOrigin === origin;
    });
    
    if (isAllowed) {
      callback(null, true);
    } else {
      console.log(`Origin ${origin} not allowed by CORS`);
      callback(new Error('Not allowed by CORS'));
    }
  },
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  credentials: true,
  maxAge: 86400 // 24 hours
}));
app.options('*', cors());
app.use(express.json({ limit: '10mb' })); // Increase limit for larger data

// Test endpoint
app.get('/', (req, res) => {
  res.send('Linesheet Generator API is running');
});

/**
 * Fetch B2B data for a company location
 * GET /api/location/:locationId/b2b-data
 */
app.get('/api/location/:locationId/b2b-data', async (req, res) => {
  try {
    const { locationId } = req.params;
    console.log(`Fetching B2B data for location: ${locationId}`);
    
    const data = await shopifyApi.fetchLocationB2BData(locationId);
    res.json(data);
  } catch (error) {
    console.error('Error fetching B2B data for location:', error);
    res.status(500).json({
      error: 'Failed to fetch B2B data',
      message: error.message
    });
  }
});

/**
 * Fetch company location data
 * GET /api/location/:locationId
 */
app.get('/api/location/:locationId', async (req, res) => {
  try {
    const { locationId } = req.params;
    console.log(`Fetching location: ${locationId}`);
    
    const locationData = await shopifyApi.fetchCompanyLocation(locationId);
    res.json(locationData);
  } catch (error) {
    console.error('Error fetching location:', error);
    res.status(500).json({
      error: 'Failed to fetch location',
      message: error.message
    });
  }
});

/**
 * Fetch catalogs for a company location
 * GET /api/location/:locationId/catalogs
 */
app.get('/api/location/:locationId/catalogs', async (req, res) => {
  try {
    const { locationId } = req.params;
    console.log(`Fetching catalogs for location: ${locationId}`);
    
    const catalogs = await shopifyApi.fetchLocationCatalogs(locationId);
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
 * GET /api/catalog/:catalogId/products
 */
app.get('/api/catalog/:catalogId/products', async (req, res) => {
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
 * Excel generation endpoint
 * POST /api/generate-linesheet
 */
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
    
    // Handle multiple files (separate linesheets)
    if (outputType === 'separate' && result.files && result.files.length > 0) {
      // For separate linesheets, we'll zip the files and send the zip
      if (result.files.length > 1) {
        // Set response headers for ZIP download - ensuring proper content type
        res.setHeader('Content-Type', 'application/zip');
        res.setHeader('Content-Disposition', `attachment; filename="${company.name.replace(/\s+/g, '_')}_Linesheets.zip"`);
        
        // Log the response headers for debugging
        console.log('Sending ZIP file with headers:', {
          'Content-Type': res.getHeader('Content-Type'),
          'Content-Disposition': res.getHeader('Content-Disposition')
        });
        
        // Create a zip stream directly to the response
        const archive = archiver('zip', {
          zlib: { level: 9 } // Compression level
        });
        
        // Listen for all archive data to be written
        archive.on('end', function() {
          console.log('Archive wrote %d bytes', archive.pointer());
        });
        
        // Handle archive warnings
        archive.on('warning', function(err) {
          if (err.code === 'ENOENT') {
            console.warn('Archive warning:', err);
          } else {
            throw err;
          }
        });
        
        // Handle archive errors
        archive.on('error', function(err) {
          console.error('Archive error:', err);
          throw err;
        });
        
        // Pipe archive data to the response
        archive.pipe(res);
        
        // Add each Excel file to the archive
        result.files.forEach((file) => {
          console.log(`Adding file to archive: ${file.filename}`);
          archive.append(file.buffer, { name: file.filename });
        });
        
        // Finalize the archive and send
        console.log('Finalizing archive...');
        archive.finalize();
      }
      // If only one file, send it directly
      else {
        const file = result.files[0];
        
        // Set response headers for Excel download
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${file.filename}"`);
        
        // Send the Excel file as the response
        res.send(file.buffer);
      }
    } 
    // Handle single file (combined linesheet)
    else {
      // Set response headers for file download
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename="${result.filename}"`);
      
      // Send the Excel file as the response
      res.send(result.buffer);
    }
    
  } catch (error) {
    console.error('Error generating Excel:', error);
    res.status(500).json({ 
      error: 'Failed to generate Excel file',
      message: error.message
    });
  }
});

// Start the server
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
  console.log(`Allowing requests from: ${allowedOrigin}`);
});