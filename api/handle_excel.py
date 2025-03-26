from http.server import BaseHTTPRequestHandler
import json
import pandas as pd
import requests
import logging
import uuid
import base64
import traceback
from datetime import datetime
from io import BytesIO
from typing import List, Dict, Any, Optional, Tuple
import argparse
import os


# Configure logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)  # Changed from INFO to DEBUG for more detailed logging

# Handler for Vercel logs
if not logger.handlers:
    handler = logging.StreamHandler()
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)

# Create a unique request ID for tracking the entire batch
REQUEST_ID = str(uuid.uuid4())
logger.info(f"Starting new request with ID: {REQUEST_ID}")

# Define the updated BookingFormData structure to match the new Excel template
# Modify the BookingFormData class to handle multiple container types
class BookingFormData:
    def __init__(
        self,
        primary_contact: str,
        contact_email: str,
        contact_phone: str,
        po_number: str,
        goods_completion_date: str,
        delivery_date: str,
        hs_code: str,
        goods_description: str,
        container_count: int,
        container_type: str,
        estimate_cargo_gross_weight: float,
        hazardous: str,
        origin_address: str,
        origin_contact: str,
        origin_phone: str,
        destination_address: str,
        destination_contact: str,
        destination_phone: str,
        special_instructions: str = None,
        # Add additional container types and counts
        container_type_2: str = None,
        container_count_2: int = None,
        container_type_3: str = None,
        container_count_3: int = None
    ):
        logger.debug(f"[{REQUEST_ID}] Creating BookingFormData object for PO: {po_number}")
        # Main fields
        self.customerCode = "AUTO"  # Default value since it's been removed from the template
        self.factoryEmail = contact_email
        self.poNumber = po_number
        
        # Addresses are now single fields in the template
        self.pickupAddress = origin_address
        self.deliveryAddress = destination_address
        
        # Extract country from address for POL/POD (simplified approach)
        # Assuming the country is the last part of the address
        logger.debug(f"[{REQUEST_ID}] Extracting country from origin address: {origin_address}")
        self.pol = self._extract_country(origin_address)
        logger.debug(f"[{REQUEST_ID}] Extracted POL: {self.pol}")
        
        logger.debug(f"[{REQUEST_ID}] Extracting country from destination address: {destination_address}")
        self.pod = self._extract_country(destination_address)
        logger.debug(f"[{REQUEST_ID}] Extracted POD: {self.pod}")
        
        # Format dates as ISO strings
        logger.debug(f"[{REQUEST_ID}] Formatting date: {goods_completion_date}")
        self.cargoReadyDateISO = self._format_date(goods_completion_date)
        logger.debug(f"[{REQUEST_ID}] Formatted cargoReadyDateISO: {self.cargoReadyDateISO}")
        
        logger.debug(f"[{REQUEST_ID}] Formatting date: {delivery_date}")
        self.goodsRequiredDateISO = self._format_date(delivery_date)
        logger.debug(f"[{REQUEST_ID}] Formatted goodsRequiredDateISO: {self.goodsRequiredDateISO}")
        
        # Container details - modified to handle multiple container types
        logger.debug(f"[{REQUEST_ID}] Setting container details - Primary Type: {container_type}, Count: {container_count}")
        
        # Build list of containers
        containers = []
        
        # Add primary container
        if container_type and container_count:
            containers.append({
                "containerType": container_type,
                "quantity": container_count
            })
        
        # Add secondary container if provided
        if container_type_2 and container_count_2:
            logger.debug(f"[{REQUEST_ID}] Adding additional container - Type: {container_type_2}, Count: {container_count_2}")
            containers.append({
                "containerType": container_type_2,
                "quantity": container_count_2
            })
        
        # Add tertiary container if provided
        if container_type_3 and container_count_3:
            logger.debug(f"[{REQUEST_ID}] Adding additional container - Type: {container_type_3}, Count: {container_count_3}")
            containers.append({
                "containerType": container_type_3,
                "quantity": container_count_3
            })
        
        self.containerDetails = {
            "containers": containers
        }
        
        # Additional fields
        self.commodityCode = hs_code
        self.incoterms = "FOB"  # Default value, could be made configurable
        self.message = special_instructions
        self.service = "quote_requested"  # Default value
        
        # Store additional contact info for potential use
        self.originContact = origin_contact
        self.originPhone = origin_phone
        self.destinationContact = destination_contact
        self.destinationPhone = destination_phone
        self.primaryContact = primary_contact
        self.contactPhone = contact_phone
        self.estimateCargoGrossWeight = estimate_cargo_gross_weight
        
        logger.debug(f"[{REQUEST_ID}] BookingFormData object created successfully for PO: {po_number}")
            
    def _extract_country(self, address: str) -> str:
        """
        Extract country from address string.
        Simplistic approach - assumes country is the last part of the address.
        """
        if not address:
            logger.warning(f"[{REQUEST_ID}] Empty address provided to _extract_country")
            return "Unknown"
            
        # Try to extract country from the last part of the address
        parts = [p.strip() for p in address.split(",")]
        logger.debug(f"[{REQUEST_ID}] Address parts: {parts}")
        
        if parts:
            country_candidates = ["USA", "US", "United States", "Canada", "Mexico", "UK", 
                                "France", "Germany", "China", "Japan", "Australia"]
            
            # Check the last few parts for a country name
            for i in range(min(3, len(parts))):
                part = parts[len(parts)-1-i]
                logger.debug(f"[{REQUEST_ID}] Checking part: {part}")
                for country in country_candidates:
                    if country.lower() in part.lower():
                        logger.debug(f"[{REQUEST_ID}] Found country match: {country}")
                        return country
            # If no known country found, return the last part
            logger.debug(f"[{REQUEST_ID}] No known country found, using last part: {parts[-1]}")
            return parts[-1]
        return "Unknown"
    
    def _format_date(self, date_str: str) -> str:
        """Convert date string to ISO format, supporting multiple formats"""
        logger.debug(f"[{REQUEST_ID}] Attempting to format date: {date_str} (type: {type(date_str)})")
        if isinstance(date_str, str):
            for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
                try:
                    logger.debug(f"[{REQUEST_ID}] Trying date format: {fmt}")
                    dt = datetime.strptime(date_str, fmt)
                    iso_format = dt.isoformat()
                    logger.debug(f"[{REQUEST_ID}] Date formatted successfully to: {iso_format}")
                    return iso_format
                except ValueError:
                    logger.debug(f"[{REQUEST_ID}] Format {fmt} failed")
                    continue
        elif isinstance(date_str, datetime):
            iso_format = date_str.isoformat()
            logger.debug(f"[{REQUEST_ID}] Date already datetime object, formatted to: {iso_format}")
            return iso_format
            
        # If we get here, return the original string and let the API validation handle it
        logger.warning(f"[{REQUEST_ID}] Could not format date: {date_str}, returning as is")
        return date_str
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization"""
        logger.debug(f"[{REQUEST_ID}] Converting BookingFormData to dictionary for PO: {self.poNumber}")
        result = {
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
            "service": self.service,
            # Added additional fields that might be needed by the API
            "contactPerson": self.primaryContact,
            "contactEmail": self.factoryEmail,
            "contactPhone": self.contactPhone,
            "estimateCargoGrossWeight": self.estimateCargoGrossWeight
        }
        logger.debug(f"[{REQUEST_ID}] Dictionary created successfully for PO: {self.poNumber}")
        return result


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
        # Log the size of the file
        logger.debug(f"[{REQUEST_ID}] File content size: {len(file_content)} bytes")
        
        # Load the workbook and select the appropriate sheet
        logger.debug(f"[{REQUEST_ID}] Creating Excel file object")
        excel_file = pd.ExcelFile(BytesIO(file_content))
        logger.debug(f"[{REQUEST_ID}] Available sheets: {excel_file.sheet_names}")
        
        sheet_name = "Orders" if "Orders" in excel_file.sheet_names else excel_file.sheet_names[0]
        logger.info(f"[{REQUEST_ID}] Using sheet: {sheet_name}")
        
        # Read the Excel file
        logger.debug(f"[{REQUEST_ID}] Reading Excel sheet: {sheet_name}")
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        logger.info(f"[{REQUEST_ID}] Successfully loaded Excel file, found {len(df)} rows")
        logger.debug(f"[{REQUEST_ID}] First few rows of data: {df.head().to_dict()}")
        
        # Skip any header rows (assuming data starts at row 3)
        if "Po Number" in df.columns or "PO Number" in df.columns:
            # Already has correct headers
            logger.info(f"[{REQUEST_ID}] Headers already correctly formatted")
            logger.debug(f"[{REQUEST_ID}] Columns: {df.columns.tolist()}")
        elif any(col for col in df.iloc[0].values if isinstance(col, str) and "po number" in col.lower()):
            # Headers are in the first row
            logger.info(f"[{REQUEST_ID}] Headers found in first row, reformatting")
            logger.debug(f"[{REQUEST_ID}] First row values: {df.iloc[0].tolist()}")
            df.columns = df.iloc[0]
            df = df.drop(0).reset_index(drop=True)
            logger.debug(f"[{REQUEST_ID}] New columns after reformatting: {df.columns.tolist()}")
        else:
            # Try to detect the header row
            logger.info(f"[{REQUEST_ID}] Searching for header row")
            header_found = False
            for i in range(min(5, len(df))):
                logger.debug(f"[{REQUEST_ID}] Checking row {i} for headers: {df.iloc[i].tolist()}")
                if any(col for col in df.iloc[i].values if isinstance(col, str) and 
                      any(term in col.lower() for term in ["po number", "primary contact"])):
                    logger.info(f"[{REQUEST_ID}] Header row found at index {i}")
                    df.columns = df.iloc[i]
                    df = df.drop(i).reset_index(drop=True)
                    header_found = True
                    logger.debug(f"[{REQUEST_ID}] New columns after finding header: {df.columns.tolist()}")
                    break
            
            if not header_found:
                logger.warning(f"[{REQUEST_ID}] Could not find header row, proceeding with existing columns")
                logger.debug(f"[{REQUEST_ID}] Using default columns: {df.columns.tolist()}")
        
        # Log the columns found
        logger.info(f"[{REQUEST_ID}] Found columns: {list(df.columns)}")
        
        # Clean up column names
        original_columns = df.columns.tolist()
        df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
        logger.debug(f"[{REQUEST_ID}] Cleaned column names: {dict(zip(original_columns, df.columns.tolist()))}")
        
        # Remove empty rows
        original_len = len(df)
        df = df.dropna(how='all')
        logger.info(f"[{REQUEST_ID}] Removed {original_len - len(df)} empty rows, {len(df)} rows remaining")
        
        # Log data types
        logger.debug(f"[{REQUEST_ID}] DataFrame data types: {df.dtypes}")
        
        # Check for important columns
        important_columns = ["Po Number", "PO Number", "Primary Contact", "Contact Email"]
        for col in important_columns:
            if col in df.columns:
                logger.info(f"[{REQUEST_ID}] Found important column: {col}")
                sample_values = df[col].head(3).tolist()
                logger.debug(f"[{REQUEST_ID}] Sample values for {col}: {sample_values}")
            else:
                logger.warning(f"[{REQUEST_ID}] Important column {col} not found")
        
        # Log memory usage
        memory_usage = df.memory_usage(deep=True).sum() / 1024
        logger.debug(f"[{REQUEST_ID}] DataFrame memory usage: {memory_usage:.2f} KB")
        
        return df
    
    except Exception as e:
        logger.error(f"[{REQUEST_ID}] Error processing Excel file: {str(e)}", exc_info=True)
        logger.debug(f"[{REQUEST_ID}] Exception traceback: {traceback.format_exc()}")
        raise Exception(f"Error processing Excel file: {str(e)}")

def create_booking_data_from_row(row: pd.Series, row_index: int) -> BookingFormData:
    """
    Create a BookingFormData object from a DataFrame row using the new template structure.
    
    Args:
        row: A pandas Series representing a row from the DataFrame
        row_index: The index of the row (for logging purposes)
    
    Returns:
        BookingFormData object
    """
    row_id = f"{REQUEST_ID}-R{row_index}"
    logger.info(f"[{row_id}] Processing row {row_index}")
    logger.debug(f"[{row_id}] Row data: {row.to_dict()}")
    
    # Extract values, handling missing data and handling column name variations
    def get_value(fields, default=""):
        """Get a value from one of several possible column names"""
        if not isinstance(fields, list):
            fields = [fields]
            
        for field in fields:
            if field in row and not pd.isna(row[field]):
                value = row[field]
                logger.debug(f"[{row_id}] Found value for field '{field}': {value}")
                return value
        
        logger.debug(f"[{row_id}] No value found for fields {fields}, using default: {default}")
        return default
    
    # Handle numeric conversions
    def safe_numeric(value, default=0):
        try:
            logger.debug(f"[{row_id}] Converting '{value}' to numeric")
            if pd.isna(value):
                logger.debug(f"[{row_id}] Value is NA, using default: {default}")
                return default
            result = float(value)
            logger.debug(f"[{row_id}] Converted to: {result}")
            return result
        except Exception as e:
            logger.warning(f"[{row_id}] Error converting '{value}' to numeric: {str(e)}")
            logger.debug(f"[{row_id}] Exception details: {traceback.format_exc()}")
            return default
    
    # Log key values and available columns
    logger.debug(f"[{row_id}] Available columns in row: {row.index.tolist()}")
    po_number = get_value(["PO Number", "Po Number", "po_number"])
    logger.info(f"[{row_id}] Processing order with PO Number: {po_number}")
    
    # Create the booking data object with the new structure
    try:
        logger.debug(f"[{row_id}] Extracting data for all fields")
        
        # Match the template column names
        primary_contact = get_value(["Primary Contact", "primary_contact"])
        contact_email = get_value(["Contact Email", "Origin Email", "contact_email"])
        contact_phone = get_value(["Contact Phone", "contact_phone"])
        
        # Date fields
        goods_completion_date = get_value(["Goods Completion Date", "goods_completion_date"])
        delivery_date = get_value(["Delivery Date", "delivery_date"])
        
        # Commodity information
        hs_code = get_value(["Commodity HS Code", "HS Code", "hs_code"])
        goods_description = get_value(["Goods Description", "goods_description"])
        
        # Container information - handle up to 3 container types from template
        container_type = get_value(["Container Type 1", "Container Type", "container_type"])
        
        # Numeric values with logging
        container_count_raw = get_value(["Container Count 1", "Container Count", "container_count"], 1)
        logger.debug(f"[{row_id}] Raw container count: {container_count_raw}")
        container_count = int(safe_numeric(container_count_raw, 1))
        logger.debug(f"[{row_id}] Processed container count: {container_count}")
        
        # Additional containers (if present)
        container_type_2 = get_value(["Container Type 2 (optional)"])
        container_count_2_raw = get_value(["Container Count 2 (optional)"])
        container_count_2 = int(safe_numeric(container_count_2_raw)) if container_count_2_raw else None
        
        container_type_3 = get_value(["Container Type 3 (optional)"])
        container_count_3_raw = get_value(["Container Count 3 (optional)"])
        container_count_3 = int(safe_numeric(container_count_3_raw)) if container_count_3_raw else None
        
        # Log additional container info if present
        if container_type_2 or container_count_2:
            logger.debug(f"[{row_id}] Additional container 2: Type={container_type_2}, Count={container_count_2}")
        if container_type_3 or container_count_3:
            logger.debug(f"[{row_id}] Additional container 3: Type={container_type_3}, Count={container_count_3}")
        
        # Weight information
        weight_raw = get_value([
            "Estimate Gross Weight per Container (optional)",
            "Estimate Cargo Gross Weight", 
            "estimate_cargo_gross_weight"
        ])
        logger.debug(f"[{row_id}] Raw cargo weight: {weight_raw}")
        weight = safe_numeric(weight_raw)
        logger.debug(f"[{row_id}] Processed cargo weight: {weight}")
        
        # Address and contact information
        origin_address = get_value(["Pickup Address", "Origin Address", "origin_address"])
        origin_contact = get_value(["Origin Contact", "origin_contact"])
        origin_phone = get_value(["Origin Phone", "origin_phone"])
        
        destination_address = get_value(["Delivery Address", "Destination Address", "destination_address"])
        destination_contact = get_value(["Destination Contact", "destination_contact"])
        destination_phone = get_value(["Destination Phone", "destination_phone"])
        
        # Port information
        pol_code = get_value(["POL (Port Code)"])
        pod_code = get_value(["POD (Port Code)"])
        
        if pol_code:
            logger.debug(f"[{row_id}] Using POL code from template: {pol_code}")
        if pod_code:
            logger.debug(f"[{row_id}] Using POD code from template: {pod_code}")
        
        # Other details
        special_instructions = get_value(["Special Instructions (optional)", "Special Instructions", "special_instructions"])
        hazardous = get_value(["Hazardous", "hazardous"], "No")
        
        # Incoterms
        incoterms = get_value(["Incoterms"])
        if incoterms:
            logger.debug(f"[{row_id}] Using incoterms from template: {incoterms}")
        
        logger.debug(f"[{row_id}] Creating BookingFormData object with extracted values")
        booking_data = BookingFormData(
            primary_contact=primary_contact,
            contact_email=contact_email,
            contact_phone=contact_phone,
            po_number=po_number,
            goods_completion_date=goods_completion_date,
            delivery_date=delivery_date,
            hs_code=hs_code,
            goods_description=goods_description,
            container_count=container_count,
            container_type=container_type,
            estimate_cargo_gross_weight=weight,
            hazardous=hazardous,
            origin_address=origin_address,
            origin_contact=origin_contact,
            origin_phone=origin_phone,
            destination_address=destination_address,
            destination_contact=destination_contact,
            destination_phone=destination_phone,
            special_instructions=special_instructions,
            # Add additional container types
            container_type_2=container_type_2,
            container_count_2=container_count_2,
            container_type_3=container_type_3,
            container_count_3=container_count_3
        )
        
        # Override POL/POD if provided directly in template
        if pol_code:
            booking_data.pol = pol_code
        if pod_code:
            booking_data.pod = pod_code
            
        # Override incoterms if provided
        if incoterms:
            booking_data.incoterms = incoterms
        
        # Log key shipping details
        total_containers = container_count
        if container_count_2:
            total_containers += container_count_2
        if container_count_3:
            total_containers += container_count_3
            
        container_summary = f"{container_type}×{container_count}"
        if container_type_2 and container_count_2:
            container_summary += f", {container_type_2}×{container_count_2}"
        if container_type_3 and container_count_3:
            container_summary += f", {container_type_3}×{container_count_3}"
            
        logger.info(f"[{row_id}] Order details - Origin: {booking_data.pol}, Destination: {booking_data.pod}, " +
                  f"Containers: {container_summary}, Total: {total_containers}")
        
        # Log all key fields
        data_dict = booking_data.to_dict()
        logger.debug(f"[{row_id}] Complete booking data object: {json.dumps(data_dict, default=str)}")
        
        return booking_data
    except Exception as e:
        logger.error(f"[{row_id}] Error creating booking data: {str(e)}", exc_info=True)
        logger.debug(f"[{row_id}] Exception traceback: {traceback.format_exc()}")
        raise Exception(f"Error creating booking data for row {row_index}: {str(e)}")

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
    logger.debug(f"[{row_id}] Validation starting for PO: {booking_data.poNumber}")
    
    # Check required fields
    required_fields = {
        "poNumber": booking_data.poNumber,
        "factoryEmail": booking_data.factoryEmail,
        "pickupAddress": booking_data.pickupAddress,
        "deliveryAddress": booking_data.deliveryAddress,
        "pol": booking_data.pol,
        "pod": booking_data.pod,
        "commodityCode": booking_data.commodityCode,
        "cargoReadyDateISO": booking_data.cargoReadyDateISO,
        "goodsRequiredDateISO": booking_data.goodsRequiredDateISO
    }
    
    # Log all field values
    for field, value in required_fields.items():
        logger.debug(f"[{row_id}] Field '{field}' value: '{value}' (type: {type(value)})")
    
    # Check for missing fields
    missing_fields = [field for field, value in required_fields.items() if not value]
    logger.debug(f"[{row_id}] Missing fields check result: {missing_fields}")
    
    if missing_fields:
        error_msg = f"Missing required fields: {', '.join(missing_fields)}"
        logger.warning(f"[{row_id}] Validation failed: {error_msg}")
        return False, error_msg
    
    # Validate email format
    logger.debug(f"[{row_id}] Validating email format: {booking_data.factoryEmail}")
    if booking_data.factoryEmail and "@" not in booking_data.factoryEmail:
        error_msg = "Invalid email format for factoryEmail"
        logger.warning(f"[{row_id}] Validation failed: {error_msg} (value: {booking_data.factoryEmail})")
        return False, error_msg
    
    # Validate container details
    logger.debug(f"[{row_id}] Validating container details: {booking_data.containerDetails}")
    if not booking_data.containerDetails or not booking_data.containerDetails.get("containers"):
        error_msg = "Missing container details"
        logger.warning(f"[{row_id}] Validation failed: {error_msg}")
        return False, error_msg
    
    # Validate container type
    container_type = booking_data.containerDetails["containers"][0].get("containerType")
    logger.debug(f"[{row_id}] Container type: {container_type}")
    if not container_type:
        error_msg = "Missing container type"
        logger.warning(f"[{row_id}] Validation failed: {error_msg}")
        return False, error_msg
    
    # Validate dates
    try:
        logger.debug(f"[{row_id}] Validating date formats")
        for date_field, date_value in {
            "cargoReadyDateISO": booking_data.cargoReadyDateISO,
            "goodsRequiredDateISO": booking_data.goodsRequiredDateISO
        }.items():
            if isinstance(date_value, str) and not date_value.startswith("20"):
                logger.warning(f"[{row_id}] Potentially invalid date format for {date_field}: {date_value}")
    except Exception as e:
        logger.debug(f"[{row_id}] Date validation error: {str(e)}")
    
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
    logger.debug(f"[{row_id}] PO Number: {booking_data.get('poNumber', 'unknown')}")
    
    try:
        # First, validate the booking data structure before sending
        logger.debug(f"[{row_id}] Pre-request validation of booking data structure")
        required_keys = ["poNumber", "factoryEmail", "pickupAddress", "deliveryAddress", "containerDetails"]
        missing_keys = [key for key in required_keys if key not in booking_data]
        if missing_keys:
            logger.warning(f"[{row_id}] Missing required keys in booking data: {missing_keys}")
        
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {auth_token[:5]}..." if auth_token else "None"  # Log only first few chars for security
        }
        logger.debug(f"[{row_id}] Request headers prepared")
        
        # Log the request details (excluding sensitive data)
        safe_booking_data = booking_data.copy()
        if "customerCode" in safe_booking_data:
            safe_booking_data["customerCode"] = f"{safe_booking_data['customerCode'][:3]}..." if safe_booking_data["customerCode"] else None
        if "factoryEmail" in safe_booking_data:
            email = safe_booking_data["factoryEmail"]
            if email and isinstance(email, str) and "@" in email:
                parts = email.split("@")
                safe_booking_data["factoryEmail"] = f"{parts[0][:3]}...@{parts[1]}"
        
        logger.info(f"[{row_id}] API Request Headers: {headers}")
        logger.debug(f"[{row_id}] API Request Body (sanitized): {json.dumps(safe_booking_data, default=str)}")
        
        # Check JSON validity before sending
        try:
            json_data = json.dumps(booking_data)
            logger.debug(f"[{row_id}] JSON serialization successful, size: {len(json_data)} bytes")
        except Exception as e:
            logger.error(f"[{row_id}] JSON serialization error: {str(e)}")
            logger.debug(f"[{row_id}] Problem data: {str(booking_data)[:200]}")
        
        # Send the request
        logger.debug(f"[{row_id}] Sending API request to {api_url}")
        start_time = datetime.now()
        try:
            response = requests.post(
                api_url,
                headers=headers,
                json=booking_data,
                timeout=30  # 30-second timeout
            )
            logger.debug(f"[{row_id}] Request sent successfully")
        except Exception as e:
            logger.error(f"[{row_id}] Error during request.post: {str(e)}")
            logger.debug(f"[{row_id}] Request exception details: {traceback.format_exc()}")
            raise
            
        end_time = datetime.now()
        duration_ms = (end_time - start_time).total_seconds() * 1000
        
        logger.info(f"[{row_id}] API Response received in {duration_ms:.2f}ms with status code: {response.status_code}")
        logger.debug(f"[{row_id}] Response headers: {dict(response.headers)}")
        
        # Check if the request was successful
        if response.status_code >= 200 and response.status_code < 300:
            logger.debug(f"[{row_id}] Successful status code: {response.status_code}")
            
            # Try to parse the JSON response
            try:
                response_data = response.json()
                logger.debug(f"[{row_id}] Response JSON parsed successfully")
            except json.JSONDecodeError as e:
                logger.error(f"[{row_id}] Failed to parse JSON response: {str(e)}")
                logger.debug(f"[{row_id}] Raw response content: {response.text[:500]}")
                response_data = {"raw_text": response.text[:500] + "..." if len(response.text) > 500 else response.text}
            
            # Log response summary (avoiding sensitive data)
            if isinstance(response_data, dict):
                shipment_id = response_data.get("data", {}).get("shipmentId", "unknown")
                success = response_data.get("success", False)
                logger.info(f"[{row_id}] API call successful: shipment_id={shipment_id}, success={success}")
                
                # Log any warnings or messages from the API
                if "warnings" in response_data:
                    logger.warning(f"[{row_id}] API returned warnings: {response_data['warnings']}")
                if "message" in response_data:
                    logger.info(f"[{row_id}] API message: {response_data['message']}")
            else:
                logger.info(f"[{row_id}] API call successful but response format unexpected (type: {type(response_data)})")
            
            logger.debug(f"[{row_id}] Creating success response object")
            return {
                "success": True,
                "status_code": response.status_code,
                "data": response_data,
                "duration_ms": duration_ms,
                "po_number": booking_data.get("poNumber", "unknown")  # Add PO number for tracking
            }
        else:
            logger.warning(f"[{row_id}] Error status code: {response.status_code}")
            
            # Handle error response
            error_data = {}
            try:
                error_data = response.json()
                logger.error(f"[{row_id}] API error response: {json.dumps(error_data)}")
                logger.debug(f"[{row_id}] Detailed error response: {error_data}")
            except json.JSONDecodeError:
                error_text = response.text[:200] + "..." if len(response.text) > 200 else response.text
                error_data = {"message": error_text}
                logger.error(f"[{row_id}] API error (non-JSON response): {error_text}")
                logger.debug(f"[{row_id}] Full error response text: {response.text}")
                
            logger.debug(f"[{row_id}] Creating error response object")
            return {
                "success": False,
                "status_code": response.status_code,
                "error": error_data,
                "duration_ms": duration_ms,
                "po_number": booking_data.get("poNumber", "unknown")  # Add PO number for tracking
            }
            
    except requests.exceptions.Timeout:
        error_msg = "API request timed out after 30 seconds"
        logger.error(f"[{row_id}] {error_msg}")
        logger.debug(f"[{row_id}] Timeout occurred when calling {api_url}")
        return {
            "success": False,
            "status_code": 408,
            "error": {"message": error_msg},
            "po_number": booking_data.get("poNumber", "unknown")
        }
    except requests.exceptions.ConnectionError as e:
        error_msg = f"Connection error: {str(e)}"
        logger.error(f"[{row_id}] {error_msg}")
        logger.debug(f"[{row_id}] Connection error details: {traceback.format_exc()}")
        return {
            "success": False,
            "status_code": 500,
            "error": {"message": error_msg},
            "po_number": booking_data.get("poNumber", "unknown")
        }
    except Exception as e:
        error_msg = f"Error processing booking: {str(e)}"
        logger.error(f"[{row_id}] {error_msg}", exc_info=True)
        logger.debug(f"[{row_id}] Detailed exception: {traceback.format_exc()}")
        return {
            "success": False,
            "status_code": 500,
            "error": {"message": error_msg},
            "po_number": booking_data.get("poNumber", "unknown")
        }
        
def main():
    parser = argparse.ArgumentParser(description='Process Excel file and submit booking data to API')
    parser.add_argument('excel_file', help='Path to the Excel file to process')
    parser.add_argument('--api-url', required=True, help='URL for the API endpoint')
    parser.add_argument('--auth-token', help='Authentication token for API access')
    args = parser.parse_args()
    
    logger.info(f"Processing Excel file: {args.excel_file}")
    
    # Check if file exists
    if not os.path.exists(args.excel_file):
        logger.error(f"Excel file not found: {args.excel_file}")
        return 1
    
    try:
        # Read the Excel file
        with open(args.excel_file, 'rb') as file:
            file_content = file.read()
            
        # Process the Excel file
        df = process_excel_file(file_content)
        logger.info(f"Successfully processed Excel file with {len(df)} rows")
        
        # Process each row
        results = []
        for index, row in df.iterrows():
            row_id = f"{REQUEST_ID}-R{index}"
            
            try:
                # Create booking data
                booking_data = create_booking_data_from_row(row, index)
                
                # Validate booking data
                is_valid, error_msg = validate_booking_data(booking_data, row_id)
                if not is_valid:
                    logger.error(f"[{row_id}] Invalid booking data: {error_msg}")
                    results.append({
                        "row": index, 
                        "po_number": booking_data.poNumber, 
                        "success": False, 
                        "error": error_msg
                    })
                    continue
                
                # Submit booking
                if args.api_url:
                    response = process_booking(booking_data.to_dict(), args.api_url, args.auth_token, row_id)
                    results.append({
                        "row": index, 
                        "po_number": booking_data.poNumber, 
                        "success": response.get("success", False),
                        "status_code": response.get("status_code"),
                        "error": response.get("error"),
                        "data": response.get("data")
                    })
                else:
                    # Dry run - just validate without submitting
                    results.append({
                        "row": index, 
                        "po_number": booking_data.poNumber, 
                        "success": True, 
                        "status": "validated but not submitted (dry run)"
                    })
            except Exception as e:
                logger.error(f"[{row_id}] Error processing row {index}: {str(e)}", exc_info=True)
                results.append({
                    "row": index, 
                    "success": False, 
                    "error": str(e)
                })
        
        # Print summary
        success_count = sum(1 for r in results if r.get("success"))
        logger.info(f"Processed {len(results)} rows: {success_count} successful, {len(results) - success_count} failed")
        
        return 0
    except Exception as e:
        logger.error(f"Error in main execution: {str(e)}", exc_info=True)
        return 1

if __name__ == "__main__":
    exit(main())