"""
Azure AD Setup Helper.

Interactive script to configure the Azure AD app registration
and verify connectivity to the SharePoint site.

Usage:
    python setup_azure.py
"""

import json
import os
import sys


def main():
    print("=" * 60)
    print("Arbitration Exhibit System — Azure AD Setup")
    print("=" * 60)
    print()
    print("This script helps configure your Azure AD app registration.")
    print("You'll need the following from your IT admin:")
    print("  1. Tenant ID (Azure AD / Entra ID)")
    print("  2. Client ID (Application ID)")
    print("  3. Client Secret (for background sync scripts)")
    print()

    # Collect values
    tenant_id = input("Enter Tenant ID: ").strip()
    client_id = input("Enter Client ID (Application ID): ").strip()
    client_secret = input("Enter Client Secret (leave blank for delegated-only): ").strip()

    if not tenant_id or not client_id:
        print("ERROR: Tenant ID and Client ID are required.")
        sys.exit(1)

    # Update settings file
    settings_path = os.path.join(os.path.dirname(__file__), "config", "settings.py")
    with open(settings_path, "r") as f:
        content = f.read()

    content = content.replace('AZURE_TENANT_ID = "YOUR_TENANT_ID"',
                              f'AZURE_TENANT_ID = "{tenant_id}"')
    content = content.replace('AZURE_CLIENT_ID = "YOUR_CLIENT_ID"',
                              f'AZURE_CLIENT_ID = "{client_id}"')
    if client_secret:
        content = content.replace('AZURE_CLIENT_SECRET = "YOUR_CLIENT_SECRET"',
                                  f'AZURE_CLIENT_SECRET = "{client_secret}"')

    with open(settings_path, "w") as f:
        f.write(content)

    print()
    print("Settings updated successfully!")
    print()

    # Test connection
    print("Testing connection to Microsoft Graph API...")
    try:
        from graph_api_client import GraphAPIClient

        client = GraphAPIClient(use_delegated_auth=True)
        user_name = client.get_current_user_name()
        print(f"  Authenticated as: {user_name}")

        site_id = client.get_site_id()
        print(f"  SharePoint site found: {site_id[:30]}...")

        drive_id = client.get_drive_id()
        print(f"  Document library found: {drive_id[:30]}...")

        print()
        print("Connection successful! Your system is ready to use.")

    except Exception as e:
        print(f"  Connection failed: {e}")
        print()
        print("Check your Azure AD configuration:")
        print("  - Ensure the app has these API permissions:")
        print("    * Files.ReadWrite.All")
        print("    * Sites.ReadWrite.All")
        print("    * User.Read")
        print("  - Ensure admin consent has been granted")
        print("  - Ensure redirect URI includes 'http://localhost'")
        print()
        print("You can re-run this script after fixing the configuration.")


if __name__ == "__main__":
    main()
