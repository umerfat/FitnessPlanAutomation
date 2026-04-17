#!/usr/bin/env python3
"""
Send a reviewed plan to the client.

Usage:
    python send_to_client.py "Shahid Khan"
    python send_to_client.py --file ~/edited_plan.docx --email client@gmail.com
"""

import argparse
import glob
import os
import sys

import config
from email_sender import send_plan_to_client
from sheets_reader import get_client_by_name


def main():
    parser = argparse.ArgumentParser(description="Send reviewed plan to client")
    parser.add_argument("name", nargs="?", help="Client name (looks up email from sheet, finds doc in output/)")
    parser.add_argument("--file", type=str, help="Custom file path (overrides auto-detection)")
    parser.add_argument("--email", type=str, help="Client email (overrides sheet lookup)")
    args = parser.parse_args()

    if not args.name and not (args.file and args.email):
        parser.error("Provide a client name, or both --file and --email.")

    # Resolve client email
    client_email = args.email
    client_name = args.name or "Client"

    if args.name and not client_email:
        client = get_client_by_name(args.name)
        if not client:
            print(f"Client '{args.name}' not found in the sheet.")
            print("Use --email to specify the email manually.")
            sys.exit(1)
        client_email = client.get("Email", "")
        client_name = client.get("Full Name", args.name)

    if not client_email:
        print("No email found for the client. Use --email to specify.")
        sys.exit(1)

    # Resolve file path
    filepath = args.file
    if not filepath:
        safe_name = client_name.replace(" ", "_").replace("/", "_")
        pattern = os.path.join(config.OUTPUT_DIR, f"{safe_name}_plan.docx")
        matches = glob.glob(pattern)
        if not matches:
            print(f"No plan document found at: {pattern}")
            print("Use --file to specify the file path.")
            sys.exit(1)
        filepath = matches[0]

    if not os.path.exists(filepath):
        print(f"File not found: {filepath}")
        sys.exit(1)

    # Confirm before sending
    print(f"Sending plan to:")
    print(f"  Client: {client_name}")
    print(f"  Email:  {client_email}")
    print(f"  File:   {filepath}")
    confirm = input("\nProceed? (y/n): ").strip().lower()
    if confirm != "y":
        print("Cancelled.")
        sys.exit(0)

    send_plan_to_client(client_name, client_email, filepath)
    print(f"\n✓ Plan sent to {client_name} at {client_email}")


if __name__ == "__main__":
    main()
