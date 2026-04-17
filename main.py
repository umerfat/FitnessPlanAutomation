#!/usr/bin/env python3
"""
Main orchestrator — processes new client form submissions.

Usage:
    python main.py              # Process all new (unprocessed) clients
    python main.py --all        # Reprocess all clients (ignores processed tracking)
    python main.py --name "X"   # Process a specific client by name
"""

import argparse
import sys

from health_calculator import compute_all_metrics
from plan_generator import generate_plan
from document_builder_v2 import build_document
from email_sender import send_plan_to_coach
from sheets_reader import get_new_clients, get_client_by_name, get_all_clients, mark_processed


def process_client(client_data: dict):
    """Full pipeline for one client."""
    name = client_data.get("Full Name", "Unknown")
    email = client_data.get("Email", "")

    print(f"\n{'='*60}")
    print(f"Processing: {name}")
    print(f"{'='*60}")

    # Step 1: Calculate health metrics
    print("  [1/4] Calculating health metrics...")
    metrics = compute_all_metrics(client_data)
    print(f"         BMI: {metrics['bmi']['bmi']} ({metrics['bmi']['category']})")
    print(f"         TDEE: {metrics['tdee']} kcal")
    print(f"         Target: {metrics['calories']['target_calories']} kcal")
    print(f"         Macros: P{metrics['macros']['protein_g']}g C{metrics['macros']['carbs_g']}g F{metrics['macros']['fats_g']}g")

    # Step 2: Generate plan via Gemini
    print("  [2/4] Generating personalized plan via Gemini...")
    plan = generate_plan(client_data, metrics)
    print("         Plan generated successfully.")

    # Step 3: Build Word document
    print("  [3/4] Building Word document...")
    filepath = build_document(client_data, metrics, plan)
    print(f"         Saved to: {filepath}")

    # Step 4: Email to coach for review
    print("  [4/4] Sending to coach for review...")
    send_plan_to_coach(name, email, filepath, metrics)
    print(f"         Sent to coach email.")

    # Mark as processed
    mark_processed(client_data)

    print(f"\n  ✓ Done! Review the plan and then run:")
    print(f'    python send_to_client.py "{name}"')

    return filepath


def main():
    parser = argparse.ArgumentParser(description="Fitness Plan Automation")
    parser.add_argument("--all", action="store_true", help="Process all clients (ignore processed tracking)")
    parser.add_argument("--name", type=str, help="Process a specific client by name")
    parser.add_argument("--no-email", action="store_true", help="Skip sending email to coach (just generate doc)")
    args = parser.parse_args()

    if args.name:
        client = get_client_by_name(args.name)
        if not client:
            print(f"Client '{args.name}' not found in the sheet.")
            sys.exit(1)
        process_single(client, skip_email=args.no_email)
    elif args.all:
        clients = get_all_clients()
        if not clients:
            print("No clients found in the sheet.")
            return
        print(f"Found {len(clients)} total clients. Processing all...")
        for client in clients:
            process_single(client, skip_email=args.no_email)
    else:
        clients = get_new_clients()
        if not clients:
            print("No new clients to process.")
            return
        print(f"Found {len(clients)} new client(s). Processing...")
        for client in clients:
            process_single(client, skip_email=args.no_email)

    print("\n" + "="*60)
    print("All done!")
    print("="*60)


def process_single(client_data: dict, skip_email: bool = False):
    """Process a single client, optionally skipping email."""
    name = client_data.get("Full Name", "Unknown")
    email = client_data.get("Email", "")

    print(f"\n{'='*60}")
    print(f"Processing: {name}")
    print(f"{'='*60}")

    # Step 1
    print("  [1/4] Calculating health metrics...")
    metrics = compute_all_metrics(client_data)
    print(f"         BMI: {metrics['bmi']['bmi']} ({metrics['bmi']['category']})")
    print(f"         TDEE: {metrics['tdee']} kcal")
    print(f"         Target: {metrics['calories']['target_calories']} kcal")
    print(f"         Macros: P{metrics['macros']['protein_g']}g C{metrics['macros']['carbs_g']}g F{metrics['macros']['fats_g']}g")

    # Step 2
    print("  [2/4] Generating personalized plan via Gemini...")
    plan = generate_plan(client_data, metrics)
    print("         Plan generated successfully.")

    # Step 3
    print("  [3/4] Building Word document...")
    filepath = build_document(client_data, metrics, plan)
    print(f"         Saved to: {filepath}")

    # Step 4
    if skip_email:
        print("  [4/4] Skipping email (--no-email flag).")
    else:
        print("  [4/4] Sending to coach for review...")
        send_plan_to_coach(name, email, filepath, metrics)
        print(f"         Sent to coach email.")

    mark_processed(client_data)
    print(f"\n  ✓ Done! Review the plan and then run:")
    print(f'    python send_to_client.py "{name}"')


if __name__ == "__main__":
    main()
