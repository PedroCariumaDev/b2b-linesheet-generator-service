// shopify-api.js - Handles Shopify API communication
const axios = require('axios');
const dotenv = require('dotenv');

// Load environment variables
dotenv.config();

// Get Shopify credentials from environment variables
const SHOP_URL = process.env.SHOPIFY_STORE_URL;
const API_TOKEN = process.env.SHOPIFY_ADMIN_API_TOKEN;

if (!SHOP_URL || !API_TOKEN) {
  console.error('Error: Shopify credentials not found in environment variables');
}

// Base URL for GraphQL API
const GRAPHQL_URL = `${SHOP_URL}/admin/api/2024-01/graphql.json`;

/**
 * Make a GraphQL request to Shopify Admin API
 * @param {string} query - GraphQL query
 * @param {Object} variables - Variables for the query
 * @returns {Promise<Object>} - Query result
 */
async function graphqlRequest(query, variables = {}) {
  try {
    const response = await axios({
      url: GRAPHQL_URL,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-Shopify-Access-Token': API_TOKEN
      },
      data: {
        query,
        variables
      }
    });

    if (response.data.errors) {
      throw new Error(`GraphQL errors: ${JSON.stringify(response.data.errors)}`);
    }

    return response.data.data;
  } catch (error) {
    console.error('Shopify GraphQL request error:', error.message);
    throw error;
  }
}

/**
 * Fetch B2B catalogs associated with a customer
 * @param {string} customerId - Shopify customer ID
 * @returns {Promise<Array>} - Array of catalogs
 */
async function fetchB2BCatalogs(customerId) {
  console.log('Fetching B2B catalogs for customer:', customerId);
  
  try {
    // Ensure customerId is a string and prepare it in GID format if needed
    customerId = String(customerId || '');
    if (!customerId.startsWith('gid://')) {
      customerId = `gid://shopify/Customer/${customerId}`;
    }
    
    // GraphQL query for fetching B2B customer data
    const customerB2BQuery = `
      query GetCustomerB2B($customerId: ID!) {
        customer(id: $customerId) {
          id
          b2b {
            company {
              id
              name
              catalogs(first: 20) {
                edges {
                  node {
                    id
                    name
                    status
                    metafields(first: 10) {
                      edges {
                        node {
                          key
                          value
                        }
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
    `;
    
    const customerResponse = await graphqlRequest(customerB2BQuery, { customerId });
    
    // Check if we got a valid response with catalogs
    if (!customerResponse?.customer?.b2b?.company?.catalogs?.edges) {
      console.warn('No B2B catalogs found or invalid response structure for customer:', customerId);
      
      // Try alternative approach: query all B2B catalogs
      return await fetchAllB2BCatalogs();
    }
    
    // Extract catalogs from the response
    const catalogEdges = customerResponse.customer.b2b.company.catalogs.edges;
    
    // Format catalog data
    return catalogEdges.map(edge => {
      const catalog = edge.node;
      
      // Extract season year from metafields if available
      let seasonYear = '';
      if (catalog.metafields?.edges) {
        const seasonYearMeta = catalog.metafields.edges.find(
          meta => meta.node.key === 'season_year' || meta.node.key === 'seasonYear'
        );
        
        if (seasonYearMeta) {
          seasonYear = seasonYearMeta.node.value;
        }
      }
      
      return {
        id: catalog.id,
        name: catalog.name,
        status: catalog.status,
        seasonYear: seasonYear
      };
    });
  } catch (error) {
    console.error('Error fetching B2B catalogs via GraphQL:', error);
    
    // Try alternative approach on error
    console.log('Attempting to fetch all B2B catalogs instead...');
    return await fetchAllB2BCatalogs();
  }
}

/**
 * Fetch all B2B catalogs in the store
 * Used as a fallback if customer-specific query fails
 */
async function fetchAllB2BCatalogs() {
  console.log('Fetching all B2B catalogs in the store');
  
  try {
    const allCatalogsQuery = `
      query GetAllB2BCatalogs {
        b2bCatalogs(first: 20) {
          edges {
            node {
              id
              name
              status
              company {
                name
              }
              metafields(first: 10) {
                edges {
                  node {
                    key
                    value
                  }
                }
              }
            }
          }
        }
      }
    `;
    
    const response = await graphqlRequest(allCatalogsQuery);
    
    if (!response?.b2bCatalogs?.edges) {
      console.warn('No B2B catalogs found or invalid response structure');
      
      // Fall back to mock data if GraphQL approach fails
      console.warn('Using mock catalogs data as fallback');
      return getMockCatalogs();
    }
    
    // Extract and format catalog data
    return response.b2bCatalogs.edges.map(edge => {
      const catalog = edge.node;
      
      // Extract season year from metafields if available
      let seasonYear = '';
      if (catalog.metafields?.edges) {
        const seasonYearMeta = catalog.metafields.edges.find(
          meta => meta.node.key === 'season_year' || meta.node.key === 'seasonYear'
        );
        
        if (seasonYearMeta) {
          seasonYear = seasonYearMeta.node.value;
        }
      }
      
      return {
        id: catalog.id,
        name: catalog.name,
        status: catalog.status,
        companyName: catalog.company?.name || '',
        seasonYear: seasonYear
      };
    });
  } catch (error) {
    console.error('Error fetching all B2B catalogs:', error);
    
    // Use mock data as a final fallback
    console.warn('Using mock catalogs data due to GraphQL errors');
    return getMockCatalogs();
  }
}

/**
 * Fetch products for a specific B2B catalog
 * @param {string} catalogId - B2B catalog ID
 * @returns {Promise<Array>} - Array of products
 */
async function fetchCatalogProducts(catalogId) {
  console.log('Fetching products for catalog via GraphQL:', catalogId);
  
  try {
    const catalogQuery = `
      query GetB2BCatalog($catalogId: ID!) {
        b2bCatalog(id: $catalogId) {
          id
          name
          status
          metafields(first: 10) {
            edges {
              node {
                key
                value
              }
            }
          }
          products(first: 100) {
            edges {
              node {
                id
                title
                handle
                featuredImage {
                  url
                }
                metafields(first: 15) {
                  edges {
                    node {
                      key
                      value
                    }
                  }
                }
                productType
                variants(first: 1) {
                  edges {
                    node {
                      id
                      price
                      compareAtPrice
                    }
                  }
                }
              }
            }
          }
        }
      }
    `;
    
    const response = await graphqlRequest(catalogQuery, { catalogId });
    
    if (!response?.b2bCatalog?.products?.edges) {
      console.warn('No products found for catalog or invalid response structure:', catalogId);
      return [];
    }
    
    // Format product data
    return response.b2bCatalog.products.edges.map(edge => {
      const product = edge.node;
      return formatProductData(product);
    });
  } catch (error) {
    console.error(`Error fetching products for catalog ${catalogId}:`, error);
    
    // Use mock data as a fallback
    console.warn('Using mock product data due to GraphQL errors');
    return getMockProducts(catalogId);
  }
}

/**
 * Format product data from GraphQL response into a consistent structure
 * @param {Object} product - Product data from GraphQL
 * @returns {Object} - Formatted product data
 */
function formatProductData(product) {
  // Helper function to extract metafield value
  const getMetaValue = (metafields, key) => {
    if (!metafields?.edges) return '';
    
    const metafield = metafields.edges.find(
      edge => edge.node.key === key
    );
    
    return metafield ? metafield.node.value : '';
  };
  
  // Get image URL
  let imageUrl = '/api/placeholder/120/120';
  if (product.featuredImage?.url) {
    imageUrl = product.featuredImage.url;
  }
  
  // Get pricing from the first variant
  let wholesalePrice = 0;
  let suggRetailPrice = 0;
  
  if (product.variants?.edges?.length > 0) {
    const variant = product.variants.edges[0].node;
    wholesalePrice = parseFloat(variant.price) || 0;
    suggRetailPrice = parseFloat(variant.compareAtPrice) || wholesalePrice * 2.5;
  }
  
  // Extract product metadata
  const styleNumber = getMetaValue(product.metafields, 'style_number');
  const color = getMetaValue(product.metafields, 'color');
  const colorCode = getMetaValue(product.metafields, 'color_code');
  const season = getMetaValue(product.metafields, 'season');
  const evergreen = getMetaValue(product.metafields, 'evergreen') || 'No';
  const countryOfOrigin = getMetaValue(product.metafields, 'country_of_origin');
  const fabrication = getMetaValue(product.metafields, 'fabrication');
  const materialComposition = getMetaValue(product.metafields, 'material_composition');
  const subcategory = getMetaValue(product.metafields, 'subcategory');
  const sizeBreak = getMetaValue(product.metafields, 'size_break') || '1';
  
  // Return normalized product structure
  return {
    id: product.id,
    name: product.title,
    image: imageUrl,
    styleNumber,
    color,
    colorCode,
    season,
    evergreen,
    countryOfOrigin,
    fabrication,
    materialComposition,
    category: product.productType || '',
    subcategory,
    sizeBreak,
    wholesalePrice,
    suggRetailPrice
  };
}

/**
 * Get mock catalog data for development/testing
 */
function getMockCatalogs() {
  return [
    {
      id: "gid://shopify/B2BCatalog/12345",
      name: "SS25 Style",
      status: "ACTIVE",
      seasonYear: "Spring/Summer 2025",
      startShip: "2025-01-15",
      completeShip: "2025-02-28"
    },
    {
      id: "gid://shopify/B2BCatalog/67890",
      name: "FW25 Style",
      status: "ACTIVE",
      seasonYear: "Fall/Winter 2025",
      startShip: "2025-07-15",
      completeShip: "2025-08-30"
    }
  ];
}

/**
 * Get mock product data for development/testing
 */
function getMockProducts(catalogId) {
  if (catalogId.includes('12345')) {
    return [
      {
        id: "gid://shopify/Product/prod1",
        image: "/api/placeholder/120/120",
        name: "Sneaker Model A",
        styleNumber: "12345",
        color: "Black",
        colorCode: "B001",
        season: "SS25",
        evergreen: "No",
        countryOfOrigin: "Vietnam",
        fabrication: "Canvas",
        materialComposition: "100% Cotton",
        category: "Shoes",
        subcategory: "Sneakers",
        sizeBreak: "1",
        wholesalePrice: 33.31,
        suggRetailPrice: 114.00
      },
      {
        id: "gid://shopify/Product/prod2",
        image: "/api/placeholder/120/120",
        name: "Sneaker Model B",
        styleNumber: "12346",
        color: "White",
        colorCode: "W001",
        season: "SS25",
        evergreen: "No",
        countryOfOrigin: "Vietnam",
        fabrication: "Canvas",
        materialComposition: "100% Cotton",
        category: "Shoes",
        subcategory: "Sneakers",
        sizeBreak: "1",
        wholesalePrice: 33.31,
        suggRetailPrice: 114.00
      }
    ];
  } else if (catalogId.includes('67890')) {
    return [
      {
        id: "gid://shopify/Product/prod4",
        image: "/api/placeholder/120/120",
        name: "Winter Boot A",
        styleNumber: "34567",
        color: "Brown",
        colorCode: "BR001",
        season: "FW25",
        evergreen: "No",
        countryOfOrigin: "Italy",
        fabrication: "Leather",
        materialComposition: "100% Leather",
        category: "Shoes",
        subcategory: "Boots",
        sizeBreak: "1",
        wholesalePrice: 62.50,
        suggRetailPrice: 169.00
      }
    ];
  }
  return [];
}

module.exports = {
  fetchB2BCatalogs,
  fetchCatalogProducts
};