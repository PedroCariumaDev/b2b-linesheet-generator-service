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
 * Fetch company location data by ID
 * @param {string} locationId - Company location ID
 * @returns {Promise<Object>} - Location data
 */
async function fetchCompanyLocation(locationId) {
  console.log('Fetching company location data:', locationId);
  
  try {
    // Ensure locationId is in GID format
    if (!String(locationId).startsWith('gid://')) {
      locationId = `gid://shopify/CompanyLocation/${locationId}`;
    }
    
    // Updated query to match Shopify's schema
    const locationQuery = `
      query GetCompanyLocation($locationId: ID!) {
        companyLocation(id: $locationId) {
          id
          name
          createdAt
          updatedAt
          currency
          company {
            id
            name
            externalId
            mainContact {
              id
              customer {
                firstName
                lastName
                email
                phone
              }
            }
          }
          shippingAddress {
            address1
            address2
            city
            province
            zip
            country
          }
        }
      }
    `;
    
    const response = await graphqlRequest(locationQuery, { locationId });
    
    if (!response?.companyLocation) {
      console.warn('Location not found:', locationId);
      throw new Error('Location not found');
    }
    
    // Format company data from location response
    const location = response.companyLocation;
    const company = location.company || {};
    
    // Extract contact info properly from the mainContact structure
    let contact = {};
    if (company.mainContact && company.mainContact.customer) {
      const customerData = company.mainContact.customer;
      contact = {
        firstName: customerData.firstName || '',
        lastName: customerData.lastName || '',
        email: customerData.email || '',
        phone: customerData.phone || ''
      };
    }
    
    const address = location.shippingAddress || {};
    
    return {
      location: {
        id: location.id,
        name: location.name,
        currency: location.currency,
        createdAt: location.createdAt,
        updatedAt: location.updatedAt,
        address: {
          address1: address.address1 || '',
          address2: address.address2 || '',
          city: address.city || '',
          province: address.province || '',
          zip: address.zip || '',
          country: address.country || ''
        }
      },
      company: {
        id: company.id || '',
        name: company.name || '',
        externalId: company.externalId || '',
        contact: contact
      }
    };
  } catch (error) {
    console.error('Error fetching company location:', error);
    throw error;
  }
}

/**
 * Fetch catalogs for a company location
 * @param {string} locationId - Company location ID
 * @returns {Promise<Array>} - Array of catalogs
 */
async function fetchLocationCatalogs(locationId) {
  console.log('Fetching catalogs for company location:', locationId);
  
  try {
    // Ensure locationId is in GID format
    if (!String(locationId).startsWith('gid://')) {
      locationId = `gid://shopify/CompanyLocation/${locationId}`;
    }
    
    // Basic catalog query with no metafields
    const catalogsQuery = `
      query GetLocationCatalogs($locationId: ID!) {
        companyLocation(id: $locationId) {
          id
          name
          catalogs(first: 20) {
            edges {
              node {
                id
                title
                status
                priceList {
                  id
                  name
                }
              }
            }
          }
        }
      }
    `;
    
    const response = await graphqlRequest(catalogsQuery, { locationId });
    
    if (!response?.companyLocation?.catalogs?.edges) {
      console.warn('No catalogs found for location:', locationId);
      return [];
    }
    
    // Format catalog data - just return basic catalog data
    return response.companyLocation.catalogs.edges.map(edge => {
      const catalog = edge.node;
      
      return {
        id: catalog.id,
        name: catalog.title,
        status: catalog.status,
        priceListId: catalog.priceList?.id,
        priceListName: catalog.priceList?.name,
        // Products will be fetched separately
        products: []
      };
    });
  } catch (error) {
    console.error('Error fetching location catalogs:', error);
    throw error;
  }
}

/**
 * Check if a company location has a specific catalog
 * @param {string} locationId - Company location ID
 * @param {string} catalogId - Catalog ID to check
 * @returns {Promise<boolean>} - Whether the location has the catalog
 */
async function checkLocationHasCatalog(locationId, catalogId) {
  try {
    // Ensure IDs are in GID format
    if (!String(locationId).startsWith('gid://')) {
      locationId = `gid://shopify/CompanyLocation/${locationId}`;
    }
    if (!String(catalogId).startsWith('gid://')) {
      catalogId = `gid://shopify/Catalog/${catalogId}`;
    }
    
    const query = `
      query CheckLocationCatalog($locationId: ID!, $catalogId: ID!) {
        companyLocation(id: $locationId) {
          inCatalog(catalogId: $catalogId)
        }
      }
    `;
    
    const response = await graphqlRequest(query, { locationId, catalogId });
    return response?.companyLocation?.inCatalog || false;
  } catch (error) {
    console.error('Error checking location catalog:', error);
    return false;
  }
}

/**
 * Fetch products for a catalog
 * @param {string} catalogId - Catalog ID
 * @returns {Promise<Array>} - Array of products
 */
async function fetchCatalogProducts(catalogId) {
  console.log('Fetching products for catalog:', catalogId);
  
  try {
    // Ensure catalogId is in GID format
    if (!String(catalogId).startsWith('gid://')) {
      catalogId = `gid://shopify/Catalog/${catalogId}`;
    }
    
    // Basic products query
    const productsQuery = `
      query GetCatalogProducts($catalogId: ID!) {
        catalog(id: $catalogId) {
          id
          title
          publication {
            products(first: 100) {
              edges {
                node {
                  id
                  title
                  handle
                  featuredImage {
                    url
                  }
                  productType
                  styleNumber: metafield(namespace: "cariuma-v2", key: "style_number") {
                    value
                  }
                  colorCode: metafield(namespace: "cariuma-v2", key: "color_code") {
                    value
                  }
                  season: metafield(namespace: "cariuma-v2", key: "season") {
                    value
                  }
                  evergreen: metafield(namespace: "cariuma-v2", key: "evergreen") {
                    value
                  }
                  fabrication: metafield(namespace: "cariuma-v2", key: "fabrication") {
                    value
                  }
                  materialComposition: metafield(namespace: "cariuma-v2", key: "material_composition") {
                    value
                  }
                  category: metafield(namespace: "cariuma-v2", key: "category") {
                    value
                  }
                  subcategory: metafield(namespace: "cariuma-v2", key: "subcategory") {
                    value
                  }
                  sizeBreak: metafield(namespace: "cariuma-v2", key: "size_break") {
                    value
                  }
                  variants(first: 1) {
                    edges {
                      node {
                        id
                        price
                        compareAtPrice
                        selectedOptions {
                          name
                          value
                        }
                        inventoryItem {
                          countryCodeOfOrigin
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
    
    const response = await graphqlRequest(productsQuery, { catalogId });
    
    if (!response?.catalog?.publication?.products?.edges) {
      console.warn('No products found for catalog:', catalogId);
      return [];
    }
    
    // Simply format the product data
    return response.catalog.publication.products.edges.map(edge => {
      const product = edge.node;
      
      // Get image URL
      let imageUrl = '/api/placeholder/120/120';
      if (product.featuredImage?.url) {
        imageUrl = product.featuredImage.url;
      }
      
      // Get fields from the first variant
      let wholesalePrice = 0;
      let suggRetailPrice = 0;
      let countryOfOrigin = '';
      
      if (product.variants?.edges?.length > 0) {
        const variant = product.variants.edges[0].node;
        wholesalePrice = parseFloat(variant.price) || 0;
        suggRetailPrice = parseFloat(variant.price) || 0;

        // Extract countryCodeOfOrigin from inventoryItem if available
        if (variant.inventoryItem && variant.inventoryItem.countryCodeOfOrigin) {
          countryOfOrigin = variant.inventoryItem.countryCodeOfOrigin;
        }
      }

      // Find the color option (if any) among the variants
      let color = '';
      if (product.variants?.edges) {
        for (const variantEdge of product.variants.edges) {
          const variant = variantEdge.node;
          if (variant.selectedOptions) {
            for (const option of variant.selectedOptions) {
              const optionNameLower = option.name.toLowerCase();
              if (optionNameLower === 'color' || optionNameLower === 'colorway') {
                color = option.value;
                break;
              }
            }
          }
          if (color) break;
        }
      }
      
      // Return simplified product structure
      return {
        id: product.id,
        name: product.title,
        image: imageUrl,
        styleNumber: product.styleNumber ? product.styleNumber.value : '',
        color: color,
        colorCode: product.colorCode ? product.colorCode.value : '',
        season: product.season ? product.season.value : '',
        evergreen: product.evergreen && product.evergreen.value.toLowerCase() === 'true' ? 'Yes' : 'No',
        countryOfOrigin: countryOfOrigin,
        fabrication: product.fabrication ? product.fabrication.value : '',
        materialComposition: product.materialComposition ? product.materialComposition.value : '',
        category: product.category ? product.category.value : '',
        subcategory: product.subcategory ? product.subcategory.value : '',
        sizeBreak: product.sizeBreak ? product.sizeBreak.value : '',
        wholesalePrice,
        suggRetailPrice
      };
    });
  } catch (error) {
    console.error('Error fetching catalog products:', error);
    throw error;
  }
}

/**
 * Alternative approach to fetch products if catalog query fails
 * @param {string} catalogId - Catalog ID
 * @returns {Promise<Array>} - Array of products
 */
async function fetchProductsAlternative(catalogId) {
  console.log('Using alternative method to fetch products for catalog:', catalogId);
  
  try {
    // Extract numeric ID from catalog ID if it's in GID format
    let catalogNumericId = catalogId;
    if (catalogId.includes('/')) {
      catalogNumericId = catalogId.split('/').pop();
    }
    
    // Try querying products with a filter for the catalog ID in metafields or tags
    const productsQuery = `
      query GetProductsByCatalog {
        products(first: 100, query: "tag:catalog-${catalogNumericId} OR metafield_key_value:catalog_id=${catalogNumericId}") {
          edges {
            node {
              id
              title
              handle
              featuredImage {
                url
              }
              productType
              tags
              styleNumber: metafield(namespace: "cariuma-v2", key: "style_number") {
                value
              }
              colorCode: metafield(namespace: "cariuma-v2", key: "color_code") {
                value
              }
              season: metafield(namespace: "cariuma-v2", key: "season") {
                value
              }
              evergreen: metafield(namespace: "cariuma-v2", key: "evergreen") {
                value
              }
              fabrication: metafield(namespace: "cariuma-v2", key: "fabrication") {
                value
              }
              materialComposition: metafield(namespace: "cariuma-v2", key: "material_composition") {
                value
              }
              category: metafield(namespace: "cariuma-v2", key: "category") {
                value
              }
              subcategory: metafield(namespace: "cariuma-v2", key: "subcategory") {
                value
              }
              sizeBreak: metafield(namespace: "cariuma-v2", key: "size_break") {
                value
              }
              variants(first: 1) {
                edges {
                  node {
                    id
                    price
                    compareAtPrice
                    selectedOptions {
                      name
                      value
                    }
                    inventoryItem {
                      countryCodeOfOrigin
                    }
                  }
                }
              }
            }
          }
        }
      }
    `;
    
    const productsResponse = await graphqlRequest(productsQuery);
    
    if (productsResponse?.products?.edges && productsResponse.products.edges.length > 0) {
      console.log(`Found ${productsResponse.products.edges.length} products using alternative query`);
      
      // Format products data from the alternative query
      return productsResponse.products.edges.map(edge => {
        const product = edge.node;
        
        // Get image URL
        let imageUrl = '/api/placeholder/120/120';
        if (product.featuredImage?.url) {
          imageUrl = product.featuredImage.url;
        }
        
        // Get pricing from the first variant
        let wholesalePrice = 0;
        let suggRetailPrice = 0;
        let countryOfOrigin = '';
        
        if (product.variants?.edges?.length > 0) {
          const variant = product.variants.edges[0].node;
          wholesalePrice = parseFloat(variant.price) || 0;
          suggRetailPrice = parseFloat(variant.compareAtPrice) || wholesalePrice;
          
          // Extract countryCodeOfOrigin from inventoryItem if available
          if (variant.inventoryItem && variant.inventoryItem.countryCodeOfOrigin) {
            countryOfOrigin = variant.inventoryItem.countryCodeOfOrigin;
          }
        }

        // Find the color option (if any) among the variants
        let color = '';
        if (product.variants?.edges) {
          for (const variantEdge of product.variants.edges) {
            const variant = variantEdge.node;
            if (variant.selectedOptions) {
              for (const option of variant.selectedOptions) {
                const optionNameLower = option.name.toLowerCase();
                if (optionNameLower === 'color' || optionNameLower === 'colorway') {
                  color = option.value;
                  break;
                }
              }
            }
            if (color) break;
          }
        }
        
        // Return simplified product structure
        return {
          id: product.id,
          name: product.title,
          image: imageUrl,
          styleNumber: product.styleNumber ? product.styleNumber.value : '',
          color: color,
          colorCode: product.colorCode ? product.colorCode.value : '',
          season: product.season ? product.season.value : '',
          evergreen: product.evergreen && product.evergreen.value.toLowerCase() === 'true' ? 'Yes' : 'No',
          countryOfOrigin: countryOfOrigin,
          fabrication: product.fabrication ? product.fabrication.value : '',
          materialComposition: product.materialComposition ? product.materialComposition.value : '',
          category: product.category ? product.category.value : product.productType || '',
          subcategory: product.subcategory ? product.subcategory.value : '',
          sizeBreak: product.sizeBreak ? product.sizeBreak.value : '',
          wholesalePrice,
          suggRetailPrice
        };
      });
    }
    
    // If no products found, return empty array
    console.warn('No products found using alternative query for catalog:', catalogId);
    return [];
  } catch (error) {
    console.error('Alternative product query failed:', error);
    throw error;
  }
}

/**
 * Fetch all catalog data for a location
 * @param {string} locationId - Company location ID
 * @returns {Promise<Object>} - Location, company and catalogs with products
 */
async function fetchLocationB2BData(locationId) {
  console.log('Fetching complete B2B data for location:', locationId);
  
  try {
    // 1. Get the location and company data
    const locationData = await fetchCompanyLocation(locationId);
    
    // 2. Get catalogs for the location
    const catalogs = await fetchLocationCatalogs(locationId);
    
    if (catalogs.length === 0) {
      console.warn('No catalogs found for location');
      return { ...locationData, catalogs: [] };
    }
    
    // 3. Just fetch products for each catalog
    for (const catalog of catalogs) {
      try {
        catalog.products = await fetchCatalogProducts(catalog.id);
        console.log(`Fetched ${catalog.products.length} products for catalog ${catalog.name}`);
      } catch (error) {
        console.error(`Error fetching products for catalog ${catalog.name}:`, error);
        // Try alternative approach if main approach fails
        try {
          catalog.products = await fetchProductsAlternative(catalog.id);
          console.log(`Fetched ${catalog.products.length} products using alternative method for catalog ${catalog.name}`);
        } catch (altError) {
          console.error(`Alternative product fetch also failed for catalog ${catalog.name}:`, altError);
          catalog.products = [];
        }
      }
    }
    
    return { ...locationData, catalogs };
  } catch (error) {
    console.error('Error fetching complete B2B data for location:', error);
    throw error;
  }
}

module.exports = {
  fetchCompanyLocation,
  fetchLocationCatalogs,
  fetchCatalogProducts,
  fetchLocationB2BData,
  checkLocationHasCatalog
};