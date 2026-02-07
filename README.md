# FACQ/EPB Converter Odoo

Convert FACQ and EPB/Vaillant installation proposal PDFs to XLSX format and optionally import them as quotations into Odoo.
Also supports direct import of products and sale orders from Excel files.

## Features

- **PDF to XLSX Conversion**: Upload a FACQ or EPB PDF and download it as an Excel file
- **Automatic PDF Type Detection**: Detects whether PDF is FACQ invoice or EPB installation proposal
- **Product Catalog Matching**: EPB PDFs are matched against a product catalog for accurate pricing
- **Odoo Integration**: Automatically import the converted data as a new quotation in Odoo's sales module
- **Excel Import**: Import products and sale orders directly from Excel files to Odoo
- **Product Linking**: Automatically link sale order lines to existing products in Odoo by product code
- **Multiple workflows**:
  1. Download XLSX only (original functionality)
  2. Download XLSX + Import to Odoo (new functionality)
  3. Import products from Excel to Odoo
  4. Import sale orders from Excel to Odoo with automatic product linking

## EPB/Vaillant PDF Processing (New)

The system now supports EPB (Energieprestatie Berekening) installation proposal PDFs from Vaillant:

### How It Works

1. **Load Product Catalog**: First, import your product catalog via `/import-products` endpoint
2. **Upload EPB PDF**: The system will:
   - Detect that it's an EPB proposal (not FACQ)
   - Extract main components (aroTHERM, VWL modules, uniSTOR, etc.)
   - Parse the Legend section for additional items
   - Filter out indicators (1a, 2b, etc.) and section headings
   - Match items to your product catalog using smart keyword matching
   - Generate Odoo Sales XLSX with proper columns and pricing

### EPB Features

- **Smart Parsing**: 
  - Extracts main Vaillant components (aroTHERM, Hydraulic modules, uniSTOR, VRC controllers)
  - Parses Legend section while filtering noise (indicators like "1a", "2b", headings)
  - Preserves real items like pumps, valves, vessels, sensors

- **Product Matching**:
  - Matches items to catalog using keyword/synonym mapping
  - Handles NL/EN variations (pump/pomp, valve/klep, etc.)
  - Supports product code variations (VWL 8.2 AS, vwl8.2as, etc.)
  - Token-based fallback matching for unlisted items

- **Deduplication**:
  - Removes duplicate products, summing quantities
  - Ensures clean output with unique product lines

- **Output Format**:
  - Generates proper Odoo Sales XLSX with 6 columns:
    1. Orderreferentie (Order Reference)
    2. Partner ID
    3. Product (Product Name from catalog)
    4. Product omschrijving (Description)
    5. Aantal (Quantity)
    6. Prijs (Price from catalog)

### EPB Usage Example

```bash
# Step 1: Import product catalog
curl -X POST "http://localhost:8000/import-products" \
  -F "file=@product_catalog.xlsx"

# Response: {"success": true, "catalog_loaded": 150, "message": "..."}

# Step 2: Upload EPB PDF for conversion
curl -X POST "http://localhost:8000/upload-xlsx" \
  -F "file=@epb_installation_proposal.pdf" \
  -o output.xlsx

# Or upload and import to Odoo
curl -X POST "http://localhost:8000/upload-and-import" \
  -F "file=@epb_installation_proposal.pdf" \
  -F "customer_name=Klant Naam" \
  -o output.xlsx
```

## Setup

### Environment Variables

For Odoo integration, set the following environment variables:

```bash
# Required for Odoo integration
ODOO_URL=https://your-odoo-instance.com
ODOO_DB=your-database-name
ODOO_USER=your-username
ODOO_PASSWORD=your-password

# Optional: Default customer name when not specified in the form
DEFAULT_CUSTOMER_NAME=FACQ Customer
```

### Run with Docker

```bash
docker build -t facq-converter .
docker run -p 8000:8000 \
  -e ODOO_URL=https://your-odoo-instance.com \
  -e ODOO_DB=your-db \
  -e ODOO_USER=your-user \
  -e ODOO_PASSWORD=your-password \
  facq-converter
```

### Run locally

```bash
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

## API Endpoints

### PDF Conversion
- `GET /` - Web interface
- `POST /upload-xlsx` - Upload PDF, download XLSX
- `POST /upload-and-import` - Upload PDF, download XLSX, and import to Odoo

### Excel Import (New)
- `POST /import-products` - Import products from Excel file (Product (product.template).xlsx format)
- `POST /import-sale-order` - Import sale order from Excel file (Verkooporder (sale.order).xlsx format)

### Other
- `GET /health` - Health check

## Usage

1. Navigate to `http://localhost:8000`
2. Choose one of two options:
   - **Download XLSX alleen**: Just convert and download the Excel file
   - **Upload en importeer in Odoo**: Convert, download, AND create a new quotation in Odoo
3. Fill in customer details (for Odoo import)
4. Upload your FACQ PDF file
5. Download the generated XLSX file

When using the Odoo import feature, a new quotation will be created in your Odoo sales module with:
- A new or existing customer record
- All product lines from the PDF
- Quantities and prices preserved

## Excel Import Usage

### Importing Products

Import products from an Excel file with the following structure:

**Product (product.template).xlsx format:**
- Required columns: Product name column (name/Product/Naam), Price column (list_price or any column with "prijs")
- Optional columns: description/omschrijving, default_code/artikelnr/SKU
- The system automatically detects column names (flexible matching)
- Example:
  ```
  name                          | description              | list_price | default_code
  aroTHERM Split plus VWL 8.2 AS| Heat pump aroTHERM       | 2500.00    | VWL-82-AS
  Hydraulic module VWL 8.2 IS   | Hydraulic module         | 800.00     | VWL-82-IS
  Circulation pump              | Pump for heating circuit | 150.00     | PUMP-001
  3-port mixing valve           | Mixing valve 3-port      | 180.00     | VALVE-001
  ```

**For EPB PDFs:**
- The product catalog is loaded into memory and used for matching EPB items
- Products are matched using smart keyword/synonym mapping
- Unmatched items are included with price 0

**API Call:**
```bash
curl -X POST "http://localhost:8000/import-products" \
  -F "file=@Product (product.template).xlsx"
```

**Response:**
```json
{
  "success": true,
  "catalog_loaded": 150,
  "message": "Productcatalogus geladen: 150 producten in geheugen",
  "odoo_import": {
    "success": true,
    "created": 10,
    "updated": 140,
    "message": "Producten geïmporteerd: 10 aangemaakt, 140 bijgewerkt"
  }
}
```

### Importing Sale Orders

Import sale orders from an Excel file with the following structure:

**Verkooporder (sale.order).xlsx format:**
- Column headers: `ArtikelNr`, `Omschrijving`, `Hoeveelheid`, `Eenheidsprijs`, `Bedrag`, `BTW`
- Example:
  ```
  ArtikelNr | Omschrijving      | Hoeveelheid | Eenheidsprijs | Bedrag | BTW
  12345     | Product A - Kraan | 2           | 150.50        | 301.00 | 21
  67890     | Product B - Wasbak| 1           | 250.00        | 250.00 | 21
  ```

**API Call:**
```bash
curl -X POST "http://localhost:8000/import-sale-order" \
  -F "file=@Verkooporder (sale.order).xlsx" \
  -F "customer_name=Test Customer" \
  -F "customer_email=test@example.com"
```

**Product Linking:**
- Sale order lines will automatically be linked to existing products in Odoo based on the `ArtikelNr` (product code)
- If a product with the matching `default_code` is found, a product line is created with the `product_id` set
- If no matching product is found, a description-only line is created with the product code and description
- This ensures proper product tracking and inventory management in Odoo
