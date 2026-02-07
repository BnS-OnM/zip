from fastapi import FastAPI, UploadFile, File, Request, Form
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse, Response
from fastapi.templating import Jinja2Templates
from typing import Optional
import os
import logging

from app.pdf_to_xlsx import facq_pdf_to_xlsx, facq_pdf_to_xlsx_and_data
from app.pdf_detector import detect_pdf_type, PDFType
from app.pdf_epb_to_odoo import epb_pdf_to_xlsx_and_data
from app.odoo import create_quotation_from_xlsx_data, import_products_from_data
from app.xlsx_import import parse_product_xlsx, parse_sale_order_xlsx, load_product_catalog_to_state, get_product_catalog
from app.logging_middleware import StructuredLoggingMiddleware

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="FACQ PDF → XLSX Converter")

# Add structured logging middleware
app.add_middleware(StructuredLoggingMiddleware)

templates = Jinja2Templates(directory="app/templates")

# Configuration
DEFAULT_CUSTOMER_NAME = os.getenv("DEFAULT_CUSTOMER_NAME", "FACQ Customer")


@app.get("/favicon.ico")
async def favicon():
    """
    Return empty response for favicon to prevent 404 errors
    """
    return Response(status_code=204)


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """
    Homepage met uploadformulier
    """
    return templates.TemplateResponse(
        "index.html",
        {"request": request}
    )


@app.post("/upload-xlsx")
async def upload_pdf_to_xlsx(file: UploadFile = File(...)):
    """
    Upload FACQ PDF → download XLSX
    """
    if not file.filename.lower().endswith(".pdf"):
        return {"error": "Upload een PDF-bestand"}

    pdf_bytes = await file.read()

    try:
        # Detect PDF type
        pdf_type = detect_pdf_type(pdf_bytes)
        logger.info(f"Detected PDF type: {pdf_type.value}")
        
        # Check if catalog is required and loaded
        if pdf_type == PDFType.EPB_VOORSTEL:
            catalog = get_product_catalog()
            if not catalog:
                return JSONResponse(
                    status_code=400,
                    content={
                        "error": "Importeer eerst products via /import-products",
                        "detail": "EPB PDFs vereisen een productcatalogus voor matching"
                    }
                )
        
        # Use appropriate converter based on PDF type
        if pdf_type == PDFType.EPB_VOORSTEL:
            xlsx_file, _ = epb_pdf_to_xlsx_and_data(pdf_bytes)
        else:
            xlsx_file = facq_pdf_to_xlsx(pdf_bytes)
            
    except Exception as e:
        return {
            "error": "Conversie mislukt",
            "detail": str(e)
        }

    return StreamingResponse(
        xlsx_file,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": "attachment; filename=facq_offerte.xlsx"
        }
    )


@app.post("/upload-and-import")
async def upload_pdf_and_import_to_odoo(
    file: UploadFile = File(...),
    customer_name: Optional[str] = Form(None),
    customer_email: Optional[str] = Form(None)
):
    """
    Upload PDF (FACQ or EPB) → create XLSX → import to Odoo as quotation
    """
    if not file.filename.lower().endswith(".pdf"):
        return JSONResponse(
            status_code=400,
            content={"error": "Upload een PDF-bestand"}
        )

    pdf_bytes = await file.read()

    try:
        # Detect PDF type
        pdf_type = detect_pdf_type(pdf_bytes)
        logger.info(f"Detected PDF type: {pdf_type.value}")
        
        # Check if catalog is required and loaded
        if pdf_type == PDFType.EPB_VOORSTEL:
            catalog = get_product_catalog()
            if not catalog:
                return JSONResponse(
                    status_code=400,
                    content={
                        "error": "Importeer eerst products via /import-products",
                        "detail": "EPB PDFs vereisen een productcatalogus voor matching"
                    }
                )
        
        # Use appropriate converter based on PDF type
        if pdf_type == PDFType.EPB_VOORSTEL:
            xlsx_file, lines_data = epb_pdf_to_xlsx_and_data(pdf_bytes)
        else:
            # Default to FACQ parser for FACQ and unknown types
            if pdf_type == PDFType.UNKNOWN:
                logger.warning("Unknown PDF type detected, falling back to FACQ parser")
            xlsx_file, lines_data = facq_pdf_to_xlsx_and_data(pdf_bytes)
            
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "error": "Conversie mislukt",
                "detail": str(e)
            }
        )

    # Import to Odoo if data was extracted and Odoo is configured
    odoo_order_id = None
    odoo_config_error = None
    if lines_data:
        try:
            odoo_order_id = create_quotation_from_xlsx_data(
                lines=lines_data,
                customer_name=customer_name or DEFAULT_CUSTOMER_NAME,
                customer_email=customer_email
            )
        except ValueError as e:
            # Odoo configuration is missing - this is expected when Odoo is not set up
            # Continue to return the XLSX file without Odoo import
            odoo_config_error = str(e)
        except Exception as e:
            # Unexpected error during Odoo import
            return JSONResponse(
                status_code=500,
                content={
                    "error": "Odoo import mislukt",
                    "detail": str(e),
                    "xlsx_generated": True
                }
            )

    # Return XLSX file with headers indicating Odoo import success
    headers = {
        "Content-Disposition": "attachment; filename=facq_offerte.xlsx",
        "X-Odoo-Order-Id": str(odoo_order_id) if odoo_order_id else "none",
    }
    
    if odoo_order_id:
        headers["X-Odoo-Import-Status"] = "success"
    elif odoo_config_error:
        headers["X-Odoo-Import-Status"] = "not_configured"
    else:
        # No lines data extracted from PDF
        headers["X-Odoo-Import-Status"] = "no_data"
    
    return StreamingResponse(
        xlsx_file,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers
    )


@app.get("/health")
async def health():
    """
    Healthcheck voor Railway / Render
    """
    return {"status": "ok"}


@app.post("/import-products")
async def import_products_from_excel(
    file: UploadFile = File(...)
):
    """
    Import products from Product (product.template).xlsx file into Odoo.
    Also loads the catalog into in-memory state for EPB PDF matching.
    
    Expected Excel format:
    - Column headers in first row
    - Columns: ArtikelNr/default_code, Naam/name, etc.
    """
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        return JSONResponse(
            status_code=400,
            content={"error": "Upload een Excel-bestand (.xlsx of .xls)"}
        )
    
    xlsx_bytes = await file.read()
    
    try:
        # Load product catalog into memory state
        catalog = load_product_catalog_to_state(xlsx_bytes)
        
        # Parse product data from Excel (for Odoo import)
        products_data = parse_product_xlsx(xlsx_bytes)
        
        if not products_data:
            return JSONResponse(
                status_code=400,
                content={
                    "error": "Geen producten gevonden in het Excel-bestand",
                    "detail": "Controleer of het bestand de juiste kolommen bevat"
                }
            )
        
        # Try to import products to Odoo (optional, may fail if Odoo not configured)
        odoo_stats = None
        odoo_error = None
        try:
            odoo_stats = import_products_from_data(products_data)
        except ValueError as e:
            # Odoo configuration error - continue without Odoo import
            odoo_error = str(e)
            logger.warning(f"Odoo import skipped: {odoo_error}")
        except Exception as e:
            # Unexpected error during Odoo import - log but continue since catalog was loaded
            odoo_error = str(e)
            logger.error(f"Odoo import failed with unexpected error: {odoo_error}", exc_info=True)
        
        # Build response
        response = {
            "success": True,
            "catalog_loaded": len(catalog),
            "message": f"Productcatalogus geladen: {len(catalog)} producten in geheugen"
        }
        
        if odoo_stats:
            response["odoo_import"] = {
                "success": True,
                "created": odoo_stats['created'],
                "updated": odoo_stats['updated'],
                "message": f"Producten geïmporteerd: {odoo_stats['created']} aangemaakt, {odoo_stats['updated']} bijgewerkt"
            }
        elif odoo_error:
            response["odoo_import"] = {
                "success": False,
                "error": odoo_error,
                "message": "Odoo import overgeslagen (configuratie ontbreekt of fout opgetreden)"
            }
        
        return JSONResponse(
            status_code=200,
            content=response
        )
        
    except Exception as e:
        # Other errors
        return JSONResponse(
            status_code=500,
            content={
                "error": "Import mislukt",
                "detail": str(e)
            }
        )


@app.post("/import-sale-order")
async def import_sale_order_from_excel(
    file: UploadFile = File(...),
    customer_name: Optional[str] = Form(None),
    customer_email: Optional[str] = Form(None)
):
    """
    Import sale order from Verkooporder (sale.order).xlsx file into Odoo.
    
    This endpoint:
    1. Parses the Excel file to extract order lines
    2. Matches products by their code to existing products in Odoo
    3. Creates a new quotation with properly linked product lines
    
    Expected Excel format:
    - Column headers in first row
    - Columns: ArtikelNr, Omschrijving, Hoeveelheid, Eenheidsprijs, BTW
    """
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        return JSONResponse(
            status_code=400,
            content={"error": "Upload een Excel-bestand (.xlsx of .xls)"}
        )
    
    xlsx_bytes = await file.read()
    
    try:
        # Parse sale order data from Excel
        order_data = parse_sale_order_xlsx(xlsx_bytes)
        
        if not order_data.get('lines'):
            return JSONResponse(
                status_code=400,
                content={
                    "error": "Geen orderregels gevonden in het Excel-bestand",
                    "detail": "Controleer of het bestand de juiste kolommen bevat"
                }
            )
        
        # Create quotation in Odoo with product linking
        odoo_order_id = create_quotation_from_xlsx_data(
            lines=order_data['lines'],
            customer_name=customer_name or DEFAULT_CUSTOMER_NAME,
            customer_email=customer_email
        )
        
        return JSONResponse(
            status_code=200,
            content={
                "success": True,
                "message": f"Verkooporder aangemaakt in Odoo (ID: {odoo_order_id})",
                "order_id": odoo_order_id,
                "lines_count": len(order_data['lines'])
            }
        )
        
    except ValueError as e:
        # Odoo configuration error
        return JSONResponse(
            status_code=400,
            content={
                "error": "Odoo configuratie ontbreekt",
                "detail": str(e)
            }
        )
    except Exception as e:
        # Other errors
        return JSONResponse(
            status_code=500,
            content={
                "error": "Import mislukt",
                "detail": str(e)
            }
        )
