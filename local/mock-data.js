// ============================================================
// MOCK DATA — mirrors your Google Sheets Catalog tab
// Edit this to match your real line items while developing.
// ============================================================

const MOCK_CONFIG = {
  companyName:     'Acme Roofing Inc.',
  companyPhone:    '(555) 000-0000',
  companyEmail:    'contact@acmeroofing.com',
  companyLicense:  'Lic # 0000000',
  roofTypes:       ['Flat', 'Tile', 'Composition', 'Shingle', 'Metal', 'Other'],
  warrantyOptions: ['5-year', '7-year', '10-year', 'Manufacturer warranty'],
  defaultWarranty: '7-year',
  gasUrl:          'http://localhost:3000',
};

const MOCK_CATALOG = [
  { item: 'permits',         description: 'Obtain all building permits & licenses required by the city', price: 0,    unit: 'included in total',        category: 'Permits & Admin' },
  { item: 'remove_flat',     description: 'Remove and dispose of tar and gravel (flat roof)',            price: 1800, unit: 'per job',                   category: 'Removal' },
  { item: 'remove_tile',     description: 'Remove and dispose of concrete tile layer',                  price: 4200, unit: 'per job',                   category: 'Removal' },
  { item: 'remove_shingle',  description: 'Remove and dispose of composition shingles',                 price: 2400, unit: 'per layer',                  category: 'Removal' },
  { item: 'osb_716',         description: 'Install APA 7/16 OSB plywood',                              price: 3750, unit: 'per job',                   category: 'Decking' },
  { item: 'peel_stick',      description: 'Install Peel-N-Stick titanium PSU30 underlayment',          price: 1800, unit: 'per job',                   category: 'Underlayment' },
  { item: 'shingle_30yr',    description: 'Install CertainTeed 30-year asphalt composition shingles',  price: 8200, unit: 'customer choice',            category: 'Roofing Material' },
  { item: 'shingle_40yr',    description: 'Install 40-year manufactured warranty composition shingles', price: 9800, unit: 'customer choice',            category: 'Roofing Material' },
  { item: 'cleanup',         description: 'Clean all debris and take to recycling center',             price: 0,    unit: 'daily + end of job',         category: 'Cleanup & Warranty' },
];

module.exports = { MOCK_CONFIG, MOCK_CATALOG };
