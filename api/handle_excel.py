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
logger.setLevel(logging.INFO)  # Changed from DEBUG to INFO to reduce logging

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
        # Main fields
        self.customerCode = "AUTO"  # Default value since it's been removed from the template
        self.factoryEmail = contact_email
        self.poNumber = po_number
        
        # Addresses are now single fields in the template
        self.pickupAddress = origin_address
        self.deliveryAddress = destination_address
        
        # Extract country from address for POL/POD (simplified approach)
        self.pol = self._extract_country(origin_address)
        self.pod = self._extract_country(destination_address)
        
        # Format dates as ISO strings
        self.cargoReadyDateISO = self._format_date(goods_completion_date)
        self.goodsRequiredDateISO = self._format_date(delivery_date)
        
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
            containers.append({
                "containerType": container_type_2,
                "quantity": container_count_2
            })
        
        # Add tertiary container if provided
        if container_type_3 and container_count_3:
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
        self.goodsDescription = goods_description
        self.hazardous = hazardous
            
    def _extract_country(self, address: str) -> str:
        """
        Extract country from address string.
        Simplistic approach - assumes country is the last part of the address.
        """
        if not address:
            return "Unknown"
        
        # Try to extract country from the last part of the address
        parts = [p.strip() for p in address.split(",")]
        
        if parts:
            country_candidates = ["USA", "US", "United States", "Canada", "Mexico", "UK", 
                                "France", "Germany", "China", "Japan", "Australia"]
            
            # Check the last few parts for a country name
            for i in range(min(3, len(parts))):
                part = parts[len(parts)-1-i]
                for country in country_candidates:
                    if country.lower() in part.lower():
                        return country
            # If no known country found, return the last part
            return parts[-1]
        return "Unknown"
    
    def _format_date(self, date_str: str) -> str:
        """Convert date string to ISO format, supporting multiple formats"""
        if isinstance(date_str, str):
            for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
                try:
                    dt = datetime.strptime(date_str, fmt)
                    iso_format = dt.isoformat()
                    return iso_format
                except ValueError:
                    continue
        elif isinstance(date_str, datetime):
            iso_format = date_str.isoformat()
            return iso_format
        
        # If we get here, return the original string and let the API validation handle it
        return date_str
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization"""
        # Create the containers object in the format required by JavaScript
        containers = {
            "containers": self.containerDetails["containers"],
            "total_booking_price": "",  # Always blank
            "price_matched": False,      # Default to False
            "price_matched_at": datetime.now().isoformat()  # Current timestamp
        }
        
        result = {
            "pol": self.pol,
            "pod": self.pod,
            "pickup_address": self.pickupAddress,
            "delivery_address": self.deliveryAddress,
            "cargo_ready_date": self.cargoReadyDateISO,
            "goods_required_date": self.goodsRequiredDateISO,
            "containers": containers,
            "commodity": self.commodityCode,
            "factory_contact_email": self.factoryEmail,
            "incoterms": self.incoterms,
            "message": self.message,
            "user_id": "",  # Always blank
            "company_code": self.customerCode,
            "po_number": self.poNumber,
            "stage": self.service or 'quote_requested'
        }
        return result
        
    def to_dict_extended(self) -> Dict[str, Any]:
        """Convert to extended dictionary for JSON serialization, including ALL collected fields"""
        basic_dict = self.to_dict()
        
        # Add all additional fields that were collected but not included in the basic payload
        extended_dict = {
            **basic_dict,
            "goods_description": getattr(self, "goodsDescription", ""),
            "origin_contact": self.originContact,
            "origin_phone": self.originPhone,
            "destination_contact": self.destinationContact,
            "destination_phone": self.destinationPhone,
            "hazardous": getattr(self, "hazardous", "No"),
            "contact_person": self.primaryContact,
            "contact_phone": self.contactPhone,
            "estimate_cargo_weight": self.estimateCargoGrossWeight,
            "booking_data": {
                "original_template": True,
                "template_version": "2.0", 
                "processed_at": datetime.now().isoformat(),
                "request_id": REQUEST_ID
            }
        }
        
        return extended_dict


def process_excel_file(file_content: bytes) -> pd.DataFrame:
    """
    Process the Excel file and return a DataFrame of the Orders sheet.
    
    Args:
        file_content: Binary content of the Excel file
    
    Returns:
        DataFrame containing the Orders sheet data
    """
    try:
        # Load the workbook and select the appropriate sheet
        excel_file = pd.ExcelFile(BytesIO(file_content))
        
        sheet_name = "Orders" if "Orders" in excel_file.sheet_names else excel_file.sheet_names[0]
        
        # Read the Excel file
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        
        # Skip any header rows (assuming data starts at row 3)
        if "Po Number" in df.columns or "PO Number" in df.columns:
            # Already has correct headers
            pass
        elif any(col for col in df.iloc[0].values if isinstance(col, str) and "po number" in col.lower()):
            # Headers are in the first row
            df.columns = df.iloc[0]
            df = df.drop(0).reset_index(drop=True)
        else:
            # Try to detect the header row
            header_found = False
            for i in range(min(5, len(df))):
                if any(col for col in df.iloc[i].values if isinstance(col, str) and 
                      any(term in col.lower() for term in ["po number", "primary contact"])):
                    df.columns = df.iloc[i]
                    df = df.drop(i).reset_index(drop=True)
                    header_found = True
                    break
        
        # Clean up column names
        df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
        
        # Remove empty rows
        df = df.dropna(how='all')
        
        return df
    
    except Exception as e:
        logger.error(f"Error processing Excel file: {str(e)}")
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
    
    # Extract values, handling missing data and handling column name variations
    def get_value(fields, default=""):
        """Get a value from one of several possible column names"""
        if not isinstance(fields, list):
            fields = [fields]
            
        for field in fields:
            if field in row and not pd.isna(row[field]):
                value = row[field]
                return value
        
        return default
    
    # Handle numeric conversions
    def safe_numeric(value, default=0):
        try:
            if pd.isna(value):
                return default
            result = float(value)
            return result
        except Exception:
            return default
    
    # Log key values and available columns
    po_number = get_value(["PO Number", "Po Number", "po_number"])
    
    # Create the booking data object with the new structure
    try:
        # Match the template column names
        primary_contact = get_value(["Primary Contact", "primary_contact", "Contact Name"])
        contact_email = get_value(["Contact Email", "Origin Email", "contact_email", "Email Address"])
        contact_phone = get_value(["Contact Phone", "contact_phone", "Phone Number"])
        
        # Date fields
        goods_completion_date = get_value(["Goods Completion Date", "goods_completion_date", "Ready Date"])
        delivery_date = get_value(["Delivery Date", "delivery_date", "Required By Date"])
        
        # Commodity information
        hs_code = get_value(["Commodity HS Code", "HS Code", "hs_code", "Harmonized Code"])
        goods_description = get_value(["Goods Description", "goods_description", "Cargo Description"])
        
        # Container information - handle up to 3 container types from template
        container_type = get_value(["Container Type 1", "Container Type", "container_type", "Equipment Type"])
        
        # Numeric values with logging
        container_count_raw = get_value(["Container Count 1", "Container Count", "container_count", "Equipment Quantity"], 1)
        container_count = int(safe_numeric(container_count_raw, 1))
        
        # Additional containers (if present)
        container_type_2 = get_value(["Container Type 2 (optional)", "Equipment Type 2"])
        container_count_2_raw = get_value(["Container Count 2 (optional)", "Equipment Quantity 2"])
        container_count_2 = int(safe_numeric(container_count_2_raw)) if container_count_2_raw else None
        
        container_type_3 = get_value(["Container Type 3 (optional)", "Equipment Type 3"])
        container_count_3_raw = get_value(["Container Count 3 (optional)", "Equipment Quantity 3"])
        container_count_3 = int(safe_numeric(container_count_3_raw)) if container_count_3_raw else None
        
        # Weight information
        weight_raw = get_value([
            "Estimate Gross Weight per Container (optional)",
            "Estimate Cargo Gross Weight", 
            "estimate_cargo_gross_weight",
            "Cargo Weight (kg)"
        ])
        weight = safe_numeric(weight_raw)
        
        # Address and contact information
        origin_address = get_value(["Pickup Address", "Origin Address", "origin_address", "Origin Location"])
        origin_contact = get_value(["Origin Contact", "origin_contact", "Origin Contact Name"])
        origin_phone = get_value(["Origin Phone", "origin_phone", "Origin Contact Phone"])
        
        destination_address = get_value(["Delivery Address", "Destination Address", "destination_address", "Destination Location"])
        destination_contact = get_value(["Destination Contact", "destination_contact", "Destination Contact Name"])
        destination_phone = get_value(["Destination Phone", "destination_phone", "Destination Contact Phone"])
        
        # Port information
        pol_code = get_value(["POL (Port Code)", "Origin Port Code", "Port of Loading"])
        pod_code = get_value(["POD (Port Code)", "Destination Port Code", "Port of Discharge"])
        
        # Other details
        special_instructions = get_value(["Special Instructions (optional)", "Special Instructions", "special_instructions", "Additional Notes"])
        hazardous = get_value(["Hazardous", "hazardous", "Dangerous Goods"], "No")
        
        # Incoterms
        incoterms = get_value(["Incoterms", "Trade Terms"])
        
        # New field: Shipping Service Type
        service_type = get_value(["Shipping Service", "Service Type", "Mode of Transport"])
        
        # New field: Booking Agent
        booking_agent = get_value(["Booking Agent", "Agent", "Freight Agent"])
        
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
            
        # Set service type if provided
        if service_type:
            booking_data.service = service_type
            
        # Store booking agent in special instructions if provided
        if booking_agent:
            original_instructions = booking_data.message or ""
            if original_instructions:
                booking_data.message = f"Booking Agent: {booking_agent}\n\n{original_instructions}"
            else:
                booking_data.message = f"Booking Agent: {booking_agent}"
        
        return booking_data
    except Exception as e:
        logger.error(f"[{row_id}] Error creating booking data for row {row_index}: {str(e)}")
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
    
    # Check for missing fields
    missing_fields = [field for field, value in required_fields.items() if not value]
    
    if missing_fields:
        error_msg = f"Missing required fields: {', '.join(missing_fields)}"
        return False, error_msg
    
    # Validate email format
    if booking_data.factoryEmail and "@" not in booking_data.factoryEmail:
        error_msg = "Invalid email format for factoryEmail"
        return False, error_msg
    
    # Validate container details
    if not booking_data.containerDetails or not booking_data.containerDetails.get("containers"):
        error_msg = "Missing container details"
        return False, error_msg
    
    # Validate container type
    container_type = booking_data.containerDetails["containers"][0].get("containerType")
    if not container_type:
        error_msg = "Missing container type"
        return False, error_msg
    
    # All validations passed
    return True, ""


def process_booking(booking_data: BookingFormData, api_url: str, auth_token: str, row_id: str) -> Dict[str, Any]:
    """
    Submit the booking data to the API.
    
    Args:
        booking_data: BookingFormData object
        api_url: URL of the API endpoint
        auth_token: Authentication token for API access
        row_id: Unique ID for the row being processed (for logging)
    
    Returns:
        Dictionary with the API response or error information
    """
    # Generate both standard and extended payload
    standard_payload = booking_data.to_dict()
    extended_payload = booking_data.to_dict_extended()
    
    # Only log the payloads (as requested)
    po_number = standard_payload.get("po_number", "unknown")
    logger.info(f"ROW {row_id} - STANDARD PAYLOAD: {json.dumps(standard_payload, default=str)}")
    logger.info(f"ROW {row_id} - EXTENDED PAYLOAD: {json.dumps(extended_payload, default=str)}")
    
    results = []
    
    # Process both payloads
    for payload_type, payload in [("standard", standard_payload), ("extended", extended_payload)]:
        try:
            # Prepare headers
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {auth_token}" if auth_token else "",
                "X-Payload-Type": payload_type
            }
            
            # Send the request
            response = requests.post(
                api_url,
                headers=headers,
                json=payload,
                timeout=30
            )
            
            # Process response
            if response.status_code >= 200 and response.status_code < 300:
                try:
                    response_data = response.json()
                except json.JSONDecodeError:
                    response_data = {"raw_text": response.text[:200]}
                
                results.append({
                    "payload_type": payload_type,
                    "success": True,
                    "status_code": response.status_code,
                    "data": response_data,
                    "po_number": po_number
                })
            else:
                # Handle error response
                try:
                    error_data = response.json()
                except json.JSONDecodeError:
                    error_data = {"message": response.text[:200]}
                    
                results.append({
                    "payload_type": payload_type,
                    "success": False,
                    "status_code": response.status_code,
                    "error": error_data,
                    "po_number": po_number
                })
                
        except Exception as e:
            results.append({
                "payload_type": payload_type,
                "success": False,
                "status_code": 500,
                "error": {"message": f"Error: {str(e)}"},
                "po_number": po_number
            })
    
    # Return combined results
    return {
        "standard": next((r for r in results if r["payload_type"] == "standard"), None),
        "extended": next((r for r in results if r["payload_type"] == "extended"), None),
        "success": any(r["success"] for r in results),
        "po_number": po_number
    }

def main():
    parser = argparse.ArgumentParser(description='Process Excel file and submit booking data to API')
    parser.add_argument('excel_file', help='Path to the Excel file to process')
    parser.add_argument('--api-url', required=True, help='URL for the API endpoint')
    parser.add_argument('--auth-token', help='Authentication token for API access')
    args = parser.parse_args()
    
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
                    logger.error(f"Invalid booking data: {error_msg}")
                    results.append({
                        "row": index, 
                        "po_number": booking_data.poNumber, 
                        "success": False, 
                        "error": error_msg
                    })
                    continue
                
                # Submit booking
                if args.api_url:
                    response = process_booking(booking_data, args.api_url, args.auth_token, row_id)
                    results.append({
                        "row": index, 
                        "po_number": response.get("po_number", "unknown"), 
                        "success": response.get("success", False),
                        "status_code": response.get("standard", {}).get("status_code"),
                        "error": response.get("standard", {}).get("error")
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
                logger.error(f"Error processing row {index}: {str(e)}")
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
        logger.error(f"Error in main execution: {str(e)}")
        return 1

if __name__ == "__main__":
    exit(main())