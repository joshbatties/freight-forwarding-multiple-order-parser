from http.server import BaseHTTPRequestHandler
import json
import pandas as pd
import requests
import logging
import uuid
import base64
from datetime import datetime
from io import BytesIO
from typing import List, Dict, Any, Optional, Tuple

# Configure logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# Handler for Vercel logs
if not logger.handlers:
    handler = logging.StreamHandler()
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)

# Create a unique request ID for tracking the entire batch
REQUEST_ID = str(uuid.uuid4())

# Define the BookingFormData structure to match the API expectations
class BookingFormData:
    def __init__(
        self,
        customer_code: str,
        primary_contact: str,
        contact_email: str,
        contact_phone: str,
        po_number: str,
        pickup_date: str,
        delivery_date: str,
        hs_code: str,
        goods_description: str,
        quantity: int,
        weight_kg: float,
        length_cm: float,
        width_cm: float,
        height_cm: float,
        hazardous: str,
        declared_value: float,
        origin_address: str,
        origin_city: str,
        origin_state: str,
        origin_country: str,
        origin_postal_code: str,
        origin_contact: str,
        origin_phone: str,
        destination_address: str,
        destination_city: str,
        destination_state: str,
        destination_country: str,
        destination_postal_code: str,
        destination_contact: str,
        destination_phone: str,
        special_instructions: str = None
    ):
        self.customerCode = customer_code
        self.factoryEmail = contact_email  # Using contact email as factory email
        self.poNumber = po_number
        
        # Combine address fields to form complete addresses
        self.pickupAddress = f"{origin_address}, {origin_city}, {origin_state}, {origin_postal_code}, {origin_country}"
        self.deliveryAddress = f"{destination_address}, {destination_city}, {destination_state}, {destination_postal_code}, {destination_country}"
        
        # Format dates as ISO strings
        self.cargoReadyDateISO = self._format_date(pickup_date)
        self.goodsRequiredDateISO = self._format_date(delivery_date)
        
        # Set POL and POD (using origin and destination countries as a simplification)
        self.pol = origin_country
        self.pod = destination_country
        
        # Container details
        self.containerDetails = {
            "containers": [
                {
                    "containerType": self._determine_container_type(length_cm, width_cm, height_cm),
                    "quantity": quantity
                }
            ]
        }
        
        # Additional fields
        self.commodityCode = hs_code
        self.incoterms = "FOB"  # Default value, could be made configurable
        self.message = special_instructions
        self.service = "quote_requested"  # Default value
        
    def _format_date(self, date_str: str) -> str:
        """Convert date string to ISO format"""
        if isinstance(date_str, str):
            try:
                dt = datetime.strptime(date_str, "%Y-%m-%d")
                return dt.isoformat()
            except ValueError:
                # Try other possible formats
                try:
                    dt = datetime.strptime(date_str, "%d/%m/%Y")
                    return dt.isoformat()
                except ValueError:
                    pass
        elif isinstance(date_str, datetime):
            return date_str.isoformat()
            
        # If we get here, return the original string and let the API validation handle it
        return date_str
    
    def _determine_container_type(self, length: float, width: float, height: float) -> str:
        """
        Determine container type based on dimensions.
        This is a simplified approach - in reality, you would have more sophisticated logic.
        """
        volume = length * width * height
        
        if volume <= 33.2 * 100**3:  # 33.2 cubic meters (20ft standard)
            return "20ft Standard"
        elif volume <= 67.7 * 100**3:  # 67.7 cubic meters (40ft standard)
            return "40ft Standard"
        elif volume <= 76.4 * 100**3:  # 76.4 cubic meters (40ft high cube)
            return "40ft High Cube"
        else:
            return "45ft High Cube"
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization"""
        return {
            "customerCode": self.customerCode,
            "factoryEmail": self.factoryEmail,
            "poNumber": self.poNumber,
            "pickupAddress": self.pickupAddress,
            "deliveryAddress": self.deliveryAddress,
            "cargoReadyDateISO": self.cargoReadyDateISO,
            "goodsRequiredDateISO": self.goodsRequiredDateISO,
            "pol": self.pol,
            "pod": self.pod,
            "containerDetails": self.containerDetails,
            "commodityCode": self.commodityCode,
            "incoterms": self.incoterms,
            "message": self.message,
            "service": self.service
        }


def process_excel_file(file_content: bytes) -> pd.DataFrame:
    """
    Process the Excel file and return a DataFrame of the Orders sheet.
    
    Args:
        file_content: Binary content of the Excel file
    
    Returns:
        DataFrame containing the Orders sheet data
    """
    logger.info(f"[{REQUEST_ID}] Starting Excel file processing")
    try:
        # Read the Excel file
        logger.info(f"[{REQUEST_ID}] Reading Excel file into pandas DataFrame")
        df = pd.read_excel(BytesIO(file_content), sheet_name="Orders")
        logger.info(f"[{REQUEST_ID}] Successfully loaded Excel file, found {len(df)} rows")
        
        # Skip any header rows (assuming data starts at row 3)
        if "Customer Code" in df.columns:
            # Already has correct headers
            logger.info(f"[{REQUEST_ID}] Headers already correctly formatted")
            pass
        elif "Customer Code" in df.iloc[0].values:
            # Headers are in the first row
            logger.info(f"[{REQUEST_ID}] Headers found in first row, reformatting")
            df.columns = df.iloc[0]
            df = df.drop(0).reset_index(drop=True)
        else:
            # Try to detect the header row
            logger.info(f"[{REQUEST_ID}] Searching for header row")
            header_found = False
            for i in range(min(5, len(df))):
                if any(col for col in df.iloc[i].values if "customer" in str(col).lower()):
                    logger.info(f"[{REQUEST_ID}] Header row found at index {i}")
                    df.columns = df.iloc[i]
                    df = df.drop(i).reset_index(drop=True)
                    header_found = True
                    break
            
            if not header_found:
                logger.warning(f"[{REQUEST_ID}] Could not find header row, proceeding with existing columns")
        
        # Log the columns found
        logger.info(f"[{REQUEST_ID}] Found columns: {list(df.columns)}")
        
        # Clean up column names
        df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
        
        # Remove empty rows
        original_len = len(df)
        df = df.dropna(how='all')
        logger.info(f"[{REQUEST_ID}] Removed {original_len - len(df)} empty rows, {len(df)} rows remaining")
        
        return df
    
    except Exception as e:
        logger.error(f"[{REQUEST_ID}] Error processing Excel file: {str(e)}", exc_info=True)
        raise Exception(f"Error processing Excel file: {str(e)}")


def create_booking_data_from_row(row: pd.Series, row_index: int) -> BookingFormData:
    """
    Create a BookingFormData object from a DataFrame row.
    
    Args:
        row: A pandas Series representing a row from the DataFrame
        row_index: The index of the row (for logging purposes)
    
    Returns:
        BookingFormData object
    """
    row_id = f"{REQUEST_ID}-R{row_index}"
    logger.info(f"[{row_id}] Processing row {row_index}")
    
    # Extract values, handling missing data
    def get_value(field, default=""):
        try:
            value = row.get(field, default)
            result = default if pd.isna(value) else value
            return result
        except Exception as e:
            logger.warning(f"[{row_id}] Error getting value for field '{field}': {str(e)}")
            return default
    
    # Handle numeric conversions
    def safe_numeric(value, default=0):
        try:
            if pd.isna(value):
                return default
            return float(value)
        except Exception as e:
            logger.warning(f"[{row_id}] Error converting '{value}' to numeric: {str(e)}")
            return default
    
    # Log key values (being careful with sensitive data)
    po_number = get_value("PO Number")
    logger.info(f"[{row_id}] Processing order with PO Number: {po_number}")
    
    # Create the booking data object
    try:
        booking_data = BookingFormData(
            customer_code=get_value("Customer Code"),
            primary_contact=get_value("Primary Contact"),
            contact_email=get_value("Contact Email"),
            contact_phone=get_value("Contact Phone"),
            po_number=po_number,
            pickup_date=get_value("Pickup Date"),
            delivery_date=get_value("Delivery Date"),
            hs_code=get_value("HS Code"),
            goods_description=get_value("Goods Description"),
            quantity=int(safe_numeric(get_value("Quantity"), 1)),
            weight_kg=safe_numeric(get_value("Weight Kg")),
            length_cm=safe_numeric(get_value("Length Cm")),
            width_cm=safe_numeric(get_value("Width Cm")),
            height_cm=safe_numeric(get_value("Height Cm")),
            hazardous=get_value("Hazardous", "No"),
            declared_value=safe_numeric(get_value("Declared Value")),
            origin_address=get_value("Origin Address"),
            origin_city=get_value("Origin City"),
            origin_state=get_value("Origin State"),
            origin_country=get_value("Origin Country"),
            origin_postal_code=get_value("Origin Postal Code"),
            origin_contact=get_value("Origin Contact"),
            origin_phone=get_value("Origin Phone"),
            destination_address=get_value("Destination Address"),
            destination_city=get_value("Destination City"),
            destination_state=get_value("Destination State"),
            destination_country=get_value("Destination Country"),
            destination_postal_code=get_value("Destination Postal Code"),
            destination_contact=get_value("Destination Contact"),
            destination_phone=get_value("Destination Phone"),
            special_instructions=get_value("Special Instructions")
        )
        
        # Log key shipping details
        logger.info(f"[{row_id}] Order details - Origin: {booking_data.pol}, Destination: {booking_data.pod}, Container: {booking_data.containerDetails['containers'][0]['containerType']}, Quantity: {booking_data.containerDetails['containers'][0]['quantity']}")
        
        return booking_data
    except Exception as e:
        logger.error(f"[{row_id}] Error creating booking data: {str(e)}", exc_info=True)
        raise


def validate_booking_data(booking_data: BookingFormData, row_id: str) -> Tuple[bool, str]:
    """
    Validate the booking data to ensure all required fields are present.
    
    Args:
        booking_data: BookingFormData object to validate
        row_id: Unique ID for the row being processed (for logging)
    
    Returns:
        Tuple of (is_valid, error_message)
    """
    logger.info(f"[{row_id}] Validating booking data")
    
    # Check required fields
    required_fields = {
        "customerCode": booking_data.customerCode,
        "factoryEmail": booking_data.factoryEmail,
        "poNumber": booking_data.poNumber,
        "pickupAddress": booking_data.pickupAddress,
        "deliveryAddress": booking_data.deliveryAddress,
        "pol": booking_data.pol,
        "pod": booking_data.pod,
        "commodityCode": booking_data.commodityCode,
        "cargoReadyDateISO": booking_data.cargoReadyDateISO,
        "goodsRequiredDateISO": booking_data.goodsRequiredDateISO
    }
    
    # Check for missing fields
    missing_fields = [field for field, value in required_fields.items() if not value]
    
    if missing_fields:
        error_msg = f"Missing required fields: {', '.join(missing_fields)}"
        logger.warning(f"[{row_id}] Validation failed: {error_msg}")
        return False, error_msg
    
    # Validate email format
    if booking_data.factoryEmail and "@" not in booking_data.factoryEmail:
        error_msg = "Invalid email format for factoryEmail"
        logger.warning(f"[{row_id}] Validation failed: {error_msg} (value: {booking_data.factoryEmail})")
        return False, error_msg
    
    # Validate container details
    if not booking_data.containerDetails or not booking_data.containerDetails.get("containers"):
        error_msg = "Missing container details"
        logger.warning(f"[{row_id}] Validation failed: {error_msg}")
        return False, error_msg
    
    # All validations passed
    logger.info(f"[{row_id}] Validation successful")
    return True, ""


def process_booking(booking_data: Dict[str, Any], api_url: str, auth_token: str, row_id: str) -> Dict[str, Any]:
    """
    Submit the booking data to the API.
    
    Args:
        booking_data: Dictionary of booking data
        api_url: URL of the API endpoint
        auth_token: Authentication token for API access
        row_id: Unique ID for the row being processed (for logging)
    
    Returns:
        Dictionary with the API response or error information
    """
    logger.info(f"[{row_id}] Submitting booking to API: {api_url}")
    
    try:
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {auth_token[:5]}..." if auth_token else "None"  # Log only first few chars for security
        }
        
        # Log the request details (excluding sensitive data)
        safe_booking_data = booking_data.copy()
        if "customerCode" in safe_booking_data:
            safe_booking_data["customerCode"] = f"{safe_booking_data['customerCode'][:3]}..." if safe_booking_data["customerCode"] else None
        
        logger.info(f"[{row_id}] API Request Headers: {headers}")
        logger.debug(f"[{row_id}] API Request Body (sanitized): {json.dumps(safe_booking_data)}")
        
        # Send the request
        start_time = datetime.now()
        response = requests.post(
            api_url,
            headers=headers,
            json=booking_data,
            timeout=30  # 30-second timeout
        )
        end_time = datetime.now()
        duration_ms = (end_time - start_time).total_seconds() * 1000
        
        logger.info(f"[{row_id}] API Response received in {duration_ms:.2f}ms with status code: {response.status_code}")
        
        # Check if the request was successful
        if response.status_code >= 200 and response.status_code < 300:
            response_data = response.json()
            
            # Log response summary (avoiding sensitive data)
            if isinstance(response_data, dict):
                shipment_id = response_data.get("data", {}).get("shipmentId", "unknown")
                success = response_data.get("success", False)
                logger.info(f"[{row_id}] API call successful: shipment_id={shipment_id}, success={success}")
            else:
                logger.info(f"[{row_id}] API call successful but response format unexpected")
            
            return {
                "success": True,
                "status_code": response.status_code,
                "data": response_data,
                "duration_ms": duration_ms
            }
        else:
            # Handle error response
            error_data = {}
            try:
                error_data = response.json()
                logger.error(f"[{row_id}] API error response: {json.dumps(error_data)}")
            except:
                error_text = response.text[:200] + "..." if len(response.text) > 200 else response.text
                error_data = {"message": error_text}
                logger.error(f"[{row_id}] API error (non-JSON response): {error_text}")
                
            return {
                "success": False,
                "status_code": response.status_code,
                "error": error_data,
                "duration_ms": duration_ms
            }
            
    except requests.exceptions.Timeout:
        error_msg = "API request timed out after 30 seconds"
        logger.error(f"[{row_id}] {error_msg}")
        return {
            "success": False,
            "status_code": 408,
            "error": {"message": error_msg}
        }
    except requests.exceptions.ConnectionError as e:
        error_msg = f"Connection error: {str(e)}"
        logger.error(f"[{row_id}] {error_msg}")
        return {
            "success": False,
            "status_code": 500,
            "error": {"message": error_msg}
        }
    except Exception as e:
        error_msg = f"Error processing booking: {str(e)}"
        logger.error(f"[{row_id}] {error_msg}", exc_info=True)
        return {
            "success": False,
            "status_code": 500,
            "error": {"message": error_msg}
        }


def batch_process_excel(
    file_content: bytes,
    api_url: str,
    auth_token: str
) -> Dict[str, Any]:
    """
    Process all orders in the Excel file and submit them to the API.
    
    Args:
        file_content: Binary content of the Excel file
        api_url: URL of the API endpoint
        auth_token: Authentication token for API access
    
    Returns:
        Dictionary with summary of processing results
    """
    batch_start_time = datetime.now()
    logger.info(f"[{REQUEST_ID}] Starting batch processing of Excel file")
    logger.info(f"[{REQUEST_ID}] API URL: {api_url}")
    logger.info(f"[{REQUEST_ID}] File size: {len(file_content)} bytes")
    
    results = {
        "request_id": REQUEST_ID,
        "total_orders": 0,
        "successful_orders": 0,
        "failed_orders": 0,
        "start_time": batch_start_time.isoformat(),
        "order_results": []
    }
    
    try:
        # Process the Excel file
        df = process_excel_file(file_content)
        
        # Process each row
        logger.info(f"[{REQUEST_ID}] Processing {len(df)} orders")
        
        for idx, row in df.iterrows():
            row_id = f"{REQUEST_ID}-R{idx}"
            row_number = idx + 3  # Adjust for 0-indexing and header rows
            po_number = row.get("PO Number", f"Row {row_number}")
            
            logger.info(f"[{row_id}] Processing row {row_number}, PO: {po_number}")
            
            order_result = {
                "row_number": row_number,
                "po_number": po_number,
                "processing_start": datetime.now().isoformat()
            }
            
            try:
                # Create booking data from row
                booking_data = create_booking_data_from_row(row, idx)
                
                # Validate booking data
                is_valid, error_message = validate_booking_data(booking_data, row_id)
                
                if not is_valid:
                    logger.warning(f"[{row_id}] Validation failed for row {row_number}: {error_message}")
                    order_result.update({
                        "success": False,
                        "error": error_message
                    })
                else:
                    # Submit to API
                    logger.info(f"[{row_id}] Validation successful, submitting to API")
                    api_result = process_booking(booking_data.to_dict(), api_url, auth_token, row_id)
                    order_result.update(api_result)
            
            except Exception as e:
                error_msg = f"Error processing row: {str(e)}"
                logger.error(f"[{row_id}] {error_msg}", exc_info=True)
                order_result.update({
                    "success": False,
                    "error": error_msg
                })
            
            # Update counts
            results["total_orders"] += 1
            if order_result.get("success", False):
                results["successful_orders"] += 1
                logger.info(f"[{row_id}] Successfully processed row {row_number}")
            else:
                results["failed_orders"] += 1
                logger.error(f"[{row_id}] Failed to process row {row_number}: {order_result.get('error', 'Unknown error')}")
            
            # Add completion time
            order_result["processing_end"] = datetime.now().isoformat()
            
            # Add to results
            results["order_results"].append(order_result)
        
        # Add batch completion time and duration
        batch_end_time = datetime.now()
        results["end_time"] = batch_end_time.isoformat()
        results["duration_seconds"] = (batch_end_time - batch_start_time).total_seconds()
        
        logger.info(f"[{REQUEST_ID}] Batch processing completed: {results['successful_orders']}/{results['total_orders']} orders successful")
        logger.info(f"[{REQUEST_ID}] Total processing time: {results['duration_seconds']:.2f} seconds")
        
        return results
    
    except Exception as e:
        error_msg = f"Error processing batch: {str(e)}"
        logger.error(f"[{REQUEST_ID}] {error_msg}", exc_info=True)
        
        # Add batch completion time and duration
        batch_end_time = datetime.now()
        results.update({
            "success": False,
            "error": error_msg,
            "end_time": batch_end_time.isoformat(),
            "duration_seconds": (batch_end_time - batch_start_time).total_seconds()
        })
        
        return results


class handler(BaseHTTPRequestHandler):
    def set_cors_headers(self):
        """Set CORS headers for cross-origin requests"""
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def do_OPTIONS(self):
        """Handle OPTIONS requests for CORS preflight"""
        self.send_response(204)  # No Content
        self.set_cors_headers()
        self.end_headers()

    def do_GET(self):
        """Handle GET requests - simple health check"""
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.set_cors_headers()
        self.end_headers()
        
        response = {"status": "ready", "message": "Excel processor is online"}
        self.wfile.write(json.dumps(response).encode())

    def do_POST(self):
        """Handle POST requests for processing Excel files"""
        global REQUEST_ID
        REQUEST_ID = str(uuid.uuid4())
        
        logger.info(f"[{REQUEST_ID}] Received POST request")
        
        # Get content length
        content_length = int(self.headers.get('Content-Length', 0))
        if content_length == 0:
            self.send_response(400)
            self.send_header('Content-type', 'application/json')
            self.set_cors_headers()
            self.end_headers()
            
            response = {"error": "Empty request body", "request_id": REQUEST_ID}
            self.wfile.write(json.dumps(response).encode())
            return
        
        # Read request body
        post_data = self.rfile.read(content_length)
        
        try:
            # Parse request body
            body = json.loads(post_data.decode('utf-8'))
            
            # Extract required parameters
            file_content_base64 = body.get("file_content_base64")
            api_url = body.get("api_url")
            auth_token = body.get("auth_token")
            
            # Validate parameters
            if not file_content_base64:
                self.send_response(400)
                self.send_header('Content-type', 'application/json')
                self.set_cors_headers()
                self.end_headers()
                
                response = {"error": "Missing file_content_base64 parameter", "request_id": REQUEST_ID}
                self.wfile.write(json.dumps(response).encode())
                return
            
            if not api_url:
                self.send_response(400)
                self.send_header('Content-type', 'application/json')
                self.set_cors_headers()
                self.end_headers()
                
                response = {"error": "Missing api_url parameter", "request_id": REQUEST_ID}
                self.wfile.write(json.dumps(response).encode())
                return
            
            if not auth_token:
                self.send_response(400)
                self.send_header('Content-type', 'application/json')
                self.set_cors_headers()
                self.end_headers()
                
                response = {"error": "Missing auth_token parameter", "request_id": REQUEST_ID}
                self.wfile.write(json.dumps(response).encode())
                return
            
            # Decode base64 file content
            try:
                file_content = base64.b64decode(file_content_base64)
            except Exception as e:
                self.send_response(400)
                self.send_header('Content-type', 'application/json')
                self.set_cors_headers()
                self.end_headers()
                
                response = {"error": f"Invalid base64 encoding: {str(e)}", "request_id": REQUEST_ID}
                self.wfile.write(json.dumps(response).encode())
                return
            
            # Process the Excel file
            logger.info(f"[{REQUEST_ID}] Processing Excel file")
            results = batch_process_excel(file_content, api_url, auth_token)
            
            # Return results
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.set_cors_headers()
            self.end_headers()
            
            self.wfile.write(json.dumps(results).encode())
            
        except json.JSONDecodeError:
            self.send_response(400)
            self.send_header('Content-type', 'application/json')
            self.set_cors_headers()
            self.end_headers()
            
            response = {"error": "Invalid JSON in request body", "request_id": REQUEST_ID}
            self.wfile.write(json.dumps(response).encode())
        
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-type', 'application/json')
            self.set_cors_headers()
            self.end_headers()
            
            logger.error(f"[{REQUEST_ID}] Unhandled exception: {str(e)}", exc_info=True)
            response = {"error": f"Internal server error: {str(e)}", "request_id": REQUEST_ID}
            self.wfile.write(json.dumps(response).encode())