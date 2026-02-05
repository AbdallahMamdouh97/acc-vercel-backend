import json
import os
import base64
from datetime import datetime
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import urllib.parse

# Your projects dictionary (keep your existing one)
PROJECTS = {
    "Component Library Project d3163287-84fc-432d-9c06-53709be1d545": "00eac9c4-b796-480b-bfe3-da6df473d190",
    "1710 - Solutions Projects": "f3506136-3dec-4345-937a-7a024dc9613c",
    # ... keep all your projects
    "0000 - ACC Learning Hub": "69397ee6-7b2b-4cc5-8d1f-009993696a91"
}

# Get credentials from environment
CLIENT_ID = os.environ.get("AUTODESK_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("AUTODESK_CLIENT_SECRET", "")
BASE_URL = "https://developer.api.autodesk.com"


def handler(request):
    # Set CORS headers - ALLOW ALL ACC DOMAINS
    cors_headers = {
        "Access-Control-Allow-Origin": "*",  # Allow all origins
        "Access-Control-Allow-Methods": "POST, GET, OPTIONS, DELETE, PUT",
        "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Requested-With, Accept, Origin",
        "Access-Control-Allow-Credentials": "true",
        "Access-Control-Max-Age": "86400",
    }

    # Handle CORS preflight request
    if request.method == "OPTIONS":
        return {
            "statusCode": 200,
            "headers": cors_headers,
            "body": ""
        }

    try:
        # Parse JSON data
        data = json.loads(request.body)

        refresh_token = data.get("refresh_token", "").strip()
        project_name = data.get("project_name", "").strip()

        if not refresh_token or not project_name:
            return {
                "statusCode": 400,
                "headers": {
                    **cors_headers,
                    "Content-Type": "application/json"
                },
                "body": json.dumps({
                    "success": False,
                    "error": "Missing refresh_token or project_name"
                })
            }

        # Get project ID
        project_id = PROJECTS.get(project_name)
        if not project_id:
            return {
                "statusCode": 400,
                "headers": {
                    **cors_headers,
                    "Content-Type": "application/json"
                },
                "body": json.dumps({
                    "success": False,
                    "error": f"Project not found: {project_name}"
                })
            }

        # 1. Get access token
        access_token = refresh_acc_token(refresh_token)

        # 2. Get planning issue types
        issue_types = get_issue_types(access_token, project_id)

        planning_id = None
        for item in issue_types.get("results", []):
            if item.get("title", "").lower() == "planning":
                planning_id = item.get("id")
                break

        if not planning_id:
            return {
                "statusCode": 400,
                "headers": {
                    **cors_headers,
                    "Content-Type": "application/json"
                },
                "body": json.dumps({
                    "success": False,
                    "error": "No planning issues found in this project"
                })
            }

        # 3. Get issues
        issues = get_issues(access_token, project_id, planning_id, limit=50)

        if not issues:
            return {
                "statusCode": 404,
                "headers": {
                    **cors_headers,
                    "Content-Type": "application/json"
                },
                "body": json.dumps({
                    "success": False,
                    "error": "No issues to export"
                })
            }

        # 4. Create Excel file
        excel_data = create_excel_file(issues, project_name)
        excel_base64 = base64.b64encode(excel_data).decode("utf-8")

        # Return success
        return {
            "statusCode": 200,
            "headers": {
                **cors_headers,
                "Content-Type": "application/json"
            },
            "body": json.dumps({
                "success": True,
                "filename": f'{project_name.replace(" ", "_")}_issues_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',
                "excel_file": excel_base64,
                "issue_count": len(issues),
                "project_name": project_name
            })
        }

    except Exception as e:
        return {
            "statusCode": 500,
            "headers": {
                **cors_headers,
                "Content-Type": "application/json"
            },
            "body": json.dumps({
                "success": False,
                "error": str(e)
            })
        }


# Helper functions
def refresh_acc_token(refresh_token):
    """Get new access token using refresh token"""
    url = f"{BASE_URL}/authentication/v2/token"
    payload = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token
    }

    response = requests.post(url, data=payload, timeout=30)
    response.raise_for_status()
    return response.json()["access_token"]


def get_issue_types(access_token, project_id):
    """Get issue types for project"""
    url = f"{BASE_URL}/construction/issues/v1/projects/{project_id}/issue-types?include=subtypes"
    headers = {"Authorization": f"Bearer {access_token}"}

    response = requests.get(url, headers=headers, timeout=30)
    response.raise_for_status()
    return response.json()


def get_issues(access_token, project_id, planning_id, limit=50):
    """Get planning issues"""
    url = f"{BASE_URL}/construction/issues/v1/projects/{project_id}/issues"
    headers = {"Authorization": f"Bearer {access_token}"}
    params = {
        "limit": limit,
        "filter[issueTypeId]": planning_id
    }

    response = requests.get(url, headers=headers, params=params, timeout=30)
    response.raise_for_status()
    return response.json().get("results", [])


def create_excel_file(issues, project_name):
    """Create Excel file from issues"""
    wb = Workbook()
    ws = wb.active

    # Title
    ws["A1"] = f"{project_name} - ACC Issue Log"
    ws["A1"].font = Font(bold=True, size=14)

    # Date
    ws["A2"] = f"Exported: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"

    # Headers
    headers = ["No.", "Title", "Description", "Status", "Created", "Updated"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=col)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")

    # Add issues
    for i, issue in enumerate(issues, start=1):
        row = 4 + i

        ws.cell(row=row, column=1, value=i)  # No.
        ws.cell(row=row, column=2, value=issue.get("title", ""))  # Title
        ws.cell(row=row, column=3, value=issue.get("description", ""))  # Description
        ws.cell(row=row, column=4, value=issue.get("status", ""))  # Status

        # Format dates
        created = issue.get("createdAt", "")
        if created:
            ws.cell(row=row, column=5, value=created[:10])  # Created date

        updated = issue.get("updatedAt", "")
        if updated:
            ws.cell(row=row, column=6, value=updated[:10])  # Updated date

    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width

    # Save to bytes
    import io
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()