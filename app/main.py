from fastapi import FastAPI, UploadFile, File, Request, Form, Query
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse, Response
from fastapi.templating import Jinja2Templates
from typing import Optional, List, Dict
from pathlib import Path
import os
import logging

from app.pdf_to_xlsx import facq_pdf_to_xlsx, facq_pdf_to_xlsx_and_data
from app.pdf_detector import detect_pdf_type, PDFType
from app.pdf_epb_to_odoo import epb_pdf_to_xlsx_and_data
from app.odoo import create_quotation_from_xlsx_data, import_products_from_data
from app.xlsx_import import parse_product_xlsx, parse_sale_order_xlsx
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


def list_directory_recursive(path: Path, max_depth: int = 10, current_depth: int = 0) -> Dict:
    """
    Recursively list directory contents similar to 'ls -R'.
    
    Args:
        path: Path to the directory to list
        max_depth: Maximum recursion depth to prevent infinite loops
        current_depth: Current recursion depth (used internally)
    
    Returns:
        Dictionary containing directory structure with files and subdirectories
    """
    result = {
        "path": str(path),
        "type": "directory",
        "contents": []
    }
    
    if current_depth >= max_depth:
        result["error"] = "Maximum depth reached"
        return result
    
    if not path.exists():
        result["error"] = "Path does not exist"
        return result
    
    if not path.is_dir():
        result["error"] = "Path is not a directory"
        return result
    
    try:
        # List all items in the directory
        for item in sorted(path.iterdir()):
            item_info = {
                "name": item.name,
                "path": str(item),
                "type": "directory" if item.is_dir() else "file"
            }
            
            # Add file size for files
            if item.is_file():
                try:
                    item_info["size"] = item.stat().st_size
                except (OSError, PermissionError):
                    item_info["size"] = None
            
            # Recursively list subdirectories
            if item.is_dir():
                try:
                    subdirectory_result = list_directory_recursive(
                        item, 
                        max_depth=max_depth, 
                        current_depth=current_depth + 1
                    )
                    item_info["contents"] = subdirectory_result.get("contents", [])
                    # Preserve error information from subdirectories
                    if "error" in subdirectory_result:
                        item_info["error"] = subdirectory_result["error"]
                except PermissionError:
                    item_info["error"] = "Permission denied"
                    item_info["contents"] = []
            
            result["contents"].append(item_info)
            
    except PermissionError:
        result["error"] = "Permission denied"
    
    return result


@app.get("/list-directory")
async def list_directory(
    path: str = Query(default=".", description="Path to list (relative or absolute)"),
    max_depth: int = Query(default=10, ge=1, le=20, description="Maximum recursion depth")
):
    """
    List directory contents recursively (similar to 'ls -R' command).
    
    Query parameters:
    - path: Directory path to list (defaults to current directory)
    - max_depth: Maximum recursion depth (1-20, defaults to 10)
    
    Returns:
    - JSON structure with recursive directory listing
    """
    try:
        # Resolve the path
        target_path = Path(path).resolve()
        
        # Security check: Ensure we're not accessing sensitive system directories
        # For a production system, you might want to restrict to a specific base directory
        sensitive_paths = [Path("/etc"), Path("/sys"), Path("/proc"), Path("/dev"), Path("/root")]
        for sensitive in sensitive_paths:
            try:
                # Use is_relative_to for proper path checking (Python 3.9+)
                if target_path.is_relative_to(sensitive):
                    return JSONResponse(
                        status_code=403,
                        content={
                            "error": "Access to system directories is forbidden",
                            "path": str(target_path)
                        }
                    )
            except (ValueError, AttributeError):
                # Fallback for Python < 3.9 or if is_relative_to is not available
                # Use resolved paths for comparison
                if str(target_path).startswith(str(sensitive.resolve()) + "/") or str(target_path) == str(sensitive.resolve()):
                    return JSONResponse(
                        status_code=403,
                        content={
                            "error": "Access to system directories is forbidden",
                            "path": str(target_path)
                        }
                    )
        
        # Perform the recursive listing
        result = list_directory_recursive(target_path, max_depth=max_depth)
        
        # Check if there's an error at the root level (no contents listed)
        if "error" in result and len(result.get("contents", [])) == 0:
            return JSONResponse(
                status_code=400,
                content=result
            )
        
        return JSONResponse(
            status_code=200,
            content=result
        )
        
    except Exception as e:
        logger.error(f"Error listing directory: {e}")
        return JSONResponse(
            status_code=500,
            content={
                "error": "Failed to list directory",
                "detail": str(e)
            }
        )


@app.post("/import-products")
async def import_products_from_excel(
    file: UploadFile = File(...)
):
    """
    Import products from Product (product.template).xlsx file into Odoo.
    
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
        # Parse product data from Excel
        products_data = parse_product_xlsx(xlsx_bytes)
        
        if not products_data:
            return JSONResponse(
                status_code=400,
                content={
                    "error": "Geen producten gevonden in het Excel-bestand",
                    "detail": "Controleer of het bestand de juiste kolommen bevat"
                }
            )
        
        # Import products to Odoo
        stats = import_products_from_data(products_data)
        
        return JSONResponse(
            status_code=200,
            content={
                "success": True,
                "message": f"Producten geïmporteerd: {stats['created']} aangemaakt, {stats['updated']} bijgewerkt",
                "stats": stats
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
