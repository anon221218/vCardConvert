import argparse
import sys
import os
from datetime import datetime, timezone
import re
import csv
import json
from typing import Union, List, Tuple
from collections import defaultdict

# Edit this if you want to change the sort order of fields/columns
def custom_column_ordering(unique_headers: List[str]) -> List[str]:

    if not custom_column_ordering.enabled:
        return unique_headers

    field_order = [
        "Is Company",
        "Organization: Name",
        "Organization: Name (Phonetic)",
        "Organization: Department",
        "Title",
        "Name: Full",
        "Name: Prefix",
        "Name: First",
        "Name: Middle",
        "Name: Last",
        "Name: Maiden",
        "Name: Suffix",
        "Name: First (Phonetic)",
        "Name: Middle (Phonetic)",
        "Name: Last (Phonetic)",
        "Nickname",
        "Phone: iPhone",
        "Phone: iPhone (Preferred)",
        "Phone: Cell",
        "Phone: Cell (Preferred)",
        "Phone: Work",
        "Phone: Work (Preferred)",
        "Phone: Home",
        "Phone: Home (Preferred)",
    ]
    
    for prefix in ["Phone", "Pager", "Fax"]:
        for selected_field in [field for field in unique_headers if field.startswith(prefix)]:
            if selected_field not in field_order:
                field_order.append(selected_field)

    field_order2 = [
        "Email: Work",
        "Email: Work (Preferred)",
        "Email: Home",
        "Email: Home (Preferred)",
    ]
    field_order.extend(field_order2)

    for prefix in ["Email", "Address", "Date", "Relationship"]:
        for selected_field in [field for field in unique_headers if field.startswith(prefix)]:
            if selected_field not in field_order2:
                field_order.append(selected_field)

    headers = []
    for field in field_order:
        if field in unique_headers:
            headers.append(field)
            unique_headers.remove(field)

    headers.extend(unique_headers)
    return headers

regex = {
    "vcard_grouped_field": re.compile(r"^item(\d+)\.(.+)$"),
    "vcard_extra_label": re.compile(r"^_\$!<([^>]+)>!\$_$"),
    "vcard_comm_type_single": re.compile(r"^type=([A-Z]+)"),
    "vcard_comm_type_compound": re.compile(r"^type=([a-zA-Z]+):(.+)$"),
    "vcard_comm_type_pref": re.compile(r"^type=([a-zA-Z]+)(?:;type=pref)?:(.+)$"),
    "apple_yearless_date": re.compile(r".*X-APPLE-OMIT-YEAR=\d+:\d+-"),
    "phone_format": re.compile(r"^(?:\+?1[-. ]?)?(?:\(?(\d{3})\)?[-. ]?(\d{3})[-. ]?(\d{4}))$")
}

def format_address(address: str) -> str:
    if format_address.enabled:
        address = ", ".join(comp.strip() for comp in address.split(";") if comp.strip())
    return address

def format_phone(phone: str) -> str:
    if format_phone.enabled:
        match = regex["phone_format"].match(phone)
        if match:
            phone = f"{match.group(1)}-{match.group(2)}-{match.group(3)}"
    return phone

def parse_vcard(vcard_file: str, debug: bool = False) -> List[List[Tuple[str, str]]]:
    try:
        with open(vcard_file, "r") as file:
            content = file.read()
    except IOError as e:
        print(f"Error reading vCard file: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        sys.exit(1)

    entries = content.split("END:VCARD")
    entries = [entry.strip() + "\nEND:VCARD" for entry in entries if entry.strip()]
    output = []
    for index, entry in enumerate(entries):
        fields = entry.splitlines()
        groups = {}
        skip_photo = False
        save = []
        unknown_name = ""
        unknown = []

        print(f"\rProcessing: {index+1} of {len(entries)}", end="")

        for field in fields:

            if field.strip().startswith("PHOTO;"):
                skip_photo = True
                continue
            if skip_photo:
                if field.startswith(" "):
                    continue
                else:
                    skip_photo = False

            field = field.strip()

            if field.startswith("BEGIN:VCARD") or field.startswith("END:VCARD"):
                continue

            if any(field.startswith(prefix) for prefix in ["VERSION:", "PRODID:", "VND-63-SENSITIVE-CONTENT-CONFIG:"]):
                continue

            if field == "X-ABShowAs:COMPANY":
                save.append(("Is Company", "X"))
                continue

            if field.startswith("N:"):
                parts = field.split(":", 1)[1].strip().split(";")
                if parts[0]: save.append(("Name: Last", parts[0]))
                if parts[1]: save.append(("Name: First", parts[1]))
                if parts[2]: save.append(("Name: Middle", parts[2]))
                if parts[3]: save.append(("Name: Prefix", parts[3]))
                if parts[4]: save.append(("Name: Suffix", parts[4]))
                continue

            if field.startswith("FN:"):
                save.append(("Name: Full", field.split(":", 1)[1].strip()))
                if debug:
                    unknown_name = f"Name (Full): {field.split(":", 1)[1].strip()}"
                continue

            if field.startswith("NICKNAME:"):
                save.append(("Name: Nickname", field.split(":", 1)[1].strip()))
                continue

            if field.startswith("X-MAIDENNAME:"):
                save.append(("Name: Maiden", field.split(":", 1)[1].strip()))
                continue

            if field.startswith("X-PHONETIC-FIRST-NAME:"):
                save.append(("Name: First (Phonetic)", field.split(":", 1)[1].strip()))
                continue

            if field.startswith("X-PHONETIC-MIDDLE-NAME:"):
                save.append(("Name: Middle (Phonetic)", field.split(":", 1)[1].strip()))
                continue

            if field.startswith("X-PHONETIC-LAST-NAME:"):
                save.append(("Name: Last (Phonetic)", field.split(":", 1)[1].strip()))
                continue

            if field.startswith("ORG:"):
                parts = field[4:].split(";")
                if parts[0]: save.append(("Organization: Name", parts[0]))
                if parts[1]: save.append(("Organization: Department", parts[1]))
                continue

            if field.startswith("X-PHONETIC-ORG:"):
                save.append(("Organization: Name (Phonetic)", field.split(":", 1)[1].strip()))
                continue

            if field.startswith("TITLE:"):
                save.append(("Organization: Title", field.split(":", 1)[1].strip()))
                continue

            if field.startswith("EMAIL;"):
                parts = field.split(";", 1)[1].strip().split(";")
                label = "";
                for entry in parts:
                    match = regex["vcard_comm_type_single"].match(entry.strip())
                    if match and match.group(1) != "INTERNET":
                        label = match.group(1).capitalize()
                for entry in parts:
                    match = regex["vcard_comm_type_compound"].match(entry.strip())
                    if match:
                        save.append((f"Email: {label}", match.group(2)))
                        if parse_vcard.preferred and match.group(1).lower() == "pref":
                            save.append((f"Email: {label} (Preferred)", match.group(2)))
                continue

            if field.startswith("TEL;"):
                parts = field.split(";", 1)[1].strip().split(";")
                labels = [];
                for entry in parts:
                    match = regex["vcard_comm_type_single"].match(entry.strip())
                    if match and match.group(1).capitalize() != "Voice":
                        labels.append(match.group(1).capitalize())
                for entry in parts:
                    match = regex["vcard_comm_type_compound"].match(entry.strip())
                    if match:
                        if match.group(1).capitalize() != "Voice":
                            labels.append(match.group(1).capitalize())
                        label_select = "";
                        if "Pager" in labels:
                            label_select = "Pager"
                        elif "Fax" in labels:
                            for label in ["Home", "Work", "Other"]:
                                if label in labels:
                                    label_select = f"Fax: {label}"
                        elif "Applewatch" in labels:
                            label_select = "Phone: Apple Watch"
                        elif "Iphone" in labels:
                            label_select = "Phone: iPhone"
                        elif "Cell" in labels:
                            label_select = "Phone: Cell"
                        else:
                            for label in ["Home", "Work", "Other", "Main"]:
                                if label in labels:
                                    label_select = f"Phone: {label}"
                        save.append((label_select, format_phone(match.group(2))))
                        if parse_vcard.preferred and match.group(1).lower() == "pref":
                            save.append((f"{label_select} (Preferred)", format_phone(match.group(2))))
                continue

            if field.startswith("ADR;"):
                match = regex["vcard_comm_type_pref"].match(field.split(";", 1)[1].strip())
                label = match.group(1).capitalize()
                save.append((f"Address: {label}", format_address(match.group(2))))
                if parse_vcard.preferred and ";type=pref:" in field:
                    save.append((f"Address: {label} (Preferred)", format_address(match.group(2))))
                continue

            if field.startswith("X-SOCIALPROFILE;"):
                parts = field.split(";", 1)[1].strip().split(":", 1)
                label = parts[0].split("=")[1].capitalize().strip()
                save.append((f"Social: {label}", parts[1].strip()))
                continue

            if field.startswith("NOTE:"):
                save.append(("Note", field.split(":", 1)[1].strip()))
                continue

            if field.startswith("URL;"):
                parts = field.split(";", 1)[1].strip().split(":", 1)
                label = parts[0].split("=")[1].capitalize().strip()
                save.append((f"URL: {label}", parts[1].strip()))
                continue

            if field.startswith("BDAY"):
                save.append((f"Date: Birthday", regex["apple_yearless_date"].sub("", field.split(";", 1)[-1].strip()).split(":", 1)[-1].strip()))
                continue

            # Prepare grouped fields (item#.@@@@@) for processing
            match = regex["vcard_grouped_field"].match(field)
            if match:
                key = int(match.group(1))
                value = match.group(2)
                if key not in groups:
                    groups[key] = []
                groups[key].append(value)
                continue

            # Catch unknown data
            save.append(("Unknown", field))
            if debug:
                unknown.append(field)

        # Process collected grouped fields
        for group in groups:
            label = ""
            data = ""
            if groups[group][1].startswith("X-ABLabel:"):
                label = groups[group][1].split("X-ABLabel:")[1]
                data = groups[group][0]
            elif groups[group][0].startswith("X-ABLabel:"):
                label = groups[group][0].split("X-ABLabel:")[1]
                data = groups[group][1]
            else:
                # International address formats
                if groups[group][0].startswith("X-ABADR:") or groups[group][1].startswith("X-ABADR:"):
                    if groups[group][1].startswith("X-ABADR:"):
                        data = groups[group][0]
                    elif groups[group][0].startswith("X-ABADR:"):
                        data = groups[group][1]
                    match = regex["vcard_comm_type_pref"].match(data.split(";", 1)[1].strip())
                    label = match.group(1).capitalize()
                    save.append((f"Address: {label}", format_address(match.group(2))))
                    if parse_vcard.preferred and ";type=pref:" in data:
                        save.append((f"Address: {label} (Preferred)", format_address(match.group(2))))
                    continue
                # Catch unknown data groups
                save.append(("Unknown", f"Item Group: {group}, Values: {groups[group]}"))
            match = regex["vcard_extra_label"].match(label)
            if match:
                label = match.group(1).strip()

            if data.startswith("EMAIL;"):
                data = data.split(";", 1)[1].strip().split(":", 1)
                save.append((f"Email: {label}", data[1].strip()))
                if parse_vcard.preferred and data[0].endswith("pref"):
                    save.append((f"Email: {label} (Preferred)", data[1].strip()))
                continue

            if data.startswith("TEL:") or data.startswith("TEL;"):
                data = data.split(":", 1)[1].strip()
                save.append((f"Phone: {label}", format_phone(data)))
                if parse_vcard.preferred and data[0].endswith("pref"):
                    save.append((f"Phone: {label} (Preferred)", format_phone(data)))
                continue

            if data.startswith("ADR:") or data.startswith("ADR;"):
                data = data.split(":", 1)[1].strip()
                save.append((f"Address: {label}", format_address(data)))
                if parse_vcard.preferred and data[0].endswith("pref"):
                    save.append((f"Address: {label} (Preferred)", format_address(data)))
                continue

            if data.startswith("URL:") or data.startswith("URL;"):
                data = data.split(":", 1)[1].strip()
                save.append((f"URL: {label}", data))
                if parse_vcard.preferred and data[0].endswith("pref"):
                    save.append((f"URL: {label} (Preferred)", data))
                continue

            if data.startswith("X-AIM") or data.startswith("X-JABBER") or data.startswith("X-MSN") or data.startswith("X-YAHOO") or data.startswith("X-ICQ"):
                continue # these entries are also listed under IMPP;X-SERVICE-TYPE=

            if data.startswith("IMPP;"):
                data = data.split(":", 1)[1].strip()
                save.append((f"IMPP: {label}", data.split(":", 1)[1].strip()))
                if parse_vcard.preferred and data[0].endswith("pref"):
                    save.append((f"IMPP: {label} (Preferred)", data))
                continue

            if data.startswith("X-ABDATE:") or data.startswith("X-ABDATE;"):
                if "X-APPLE-OMIT-YEAR" in data:
                    save.append((f"Date: {label}", regex["apple_yearless_date"].sub("", data).strip()))
                else:
                    save.append((f"Date: {label}", data.split(":", 1)[1].strip()))
                continue

            if data.startswith("X-ABRELATEDNAMES:") or data.startswith("X-ABRELATEDNAMES;"):
                save.append((f"Relationship: {label}", data.split(":", 1)[1].strip()))
                continue

            save.append(("Unknown", f"Item Group: {group}, Values: {groups[group]}"))
            if debug:
                unknown.append(f"Item Group: {group}, Values: {groups[group]}")

        output.append(save)

        if debug and unknown:
            print(f"\nvCard Entry: {unknown_name}")
            for item in unknown:
                print(f"  {item}")

    print(f"\rProcessing: Complete               ")

    return output

def save_to_csv(save: List[List[Tuple[str, str]]], csv_file: str) -> None:
    if not save:
        return

    unique_headers = sorted(set(label for sublist in save for label, _ in sublist))
    headers = custom_column_ordering(unique_headers)

    rows = []
    for sublist in save:
        row_dict = defaultdict(list)
        for label, value in sublist:
            if label in headers:
                row_dict[label].append(value.replace("\\n", "\r\n").replace("\"", "\"\""))
        combined_row_dict = {header: ", ".join(row_dict[header]) for header in headers}
        rows.append(combined_row_dict)

    try:
        with open(csv_file, mode="w", newline="", encoding="utf-8") as file:
            writer = csv.DictWriter(file, fieldnames=headers, quoting=csv.QUOTE_ALL)
            writer.writeheader()
            writer.writerows(rows)
    except IOError as e:
        print(f"Error writing to CSV file: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

def convert_single_item_lists(data: Union[list, dict]) -> Union[list, dict]:
    if isinstance(data, list):
        data = [item for item in data if item]
        if len(data) == 1:
            return data[0]
        return data
    if isinstance(data, dict):
        for key in list(data.keys()):
            value = data[key]
            cleaned_value = convert_single_item_lists(value)
            if cleaned_value:
                data[key] = cleaned_value
            else:
                del data[key]
    return data

def save_to_json(save: List[List[Tuple[str, str]]], json_file: str) -> None:
    if not save:
        return
        
    unique_headers = sorted(set(label for sublist in save for label, _ in sublist))
    headers = custom_column_ordering(unique_headers)

    rows = []
    for sublist in save:

        row_dict = defaultdict(list)
        for label, value in sublist:
            if label in headers:
                row_dict[label].append(value.replace("\\n", "\n").replace("\"", "\\\""))

        combined_row_dict = {}

        for header in headers:
            if ":" in header:
                main_label, sub_label = header.split(":", 1)
                main_label = main_label.strip()
                sub_label = sub_label.strip()

                if main_label not in combined_row_dict:
                    combined_row_dict[main_label] = {}

                combined_row_dict[main_label].setdefault(sub_label, []).extend(row_dict[header])
            elif row_dict[header]:
                if len(row_dict[header]) == 1:
                    combined_row_dict[header] = row_dict[header][0]
                else:
                    combined_row_dict[header] = list(set(row_dict[header]))

        if combined_row_dict:
            rows.append(combined_row_dict)

    for row in rows:
        convert_single_item_lists(row)

    try:
        with open(json_file, mode="w", encoding="utf-8") as file:
            json.dump(rows, file, indent=4)
    except IOError as e:
        print(f"Error writing to JSON file: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

def print_to_console(save: List[List[Tuple[str, str]]]) -> None:

    unique_headers = sorted(set(label for sublist in save for label, _ in sublist))
    headers = custom_column_ordering(unique_headers)

    for entry in save:
        print("New vCard Entry:")
        entry_dict = dict(entry)
        for label in headers:
            value = entry_dict.get(label, None)
            if label in entry_dict and value:
                print(f"  {label}: {value}")

def main() -> None:
    parser = argparse.ArgumentParser(description="Parse/Convert vCard files to CSV/JSON (note: script designed for Apple Contacts vCard export format, values not present in Apple format files will show as unknown fields).")
    parser.add_argument("input", type=str, help="Input file (required), the vCard file to be parsed")
    parser.add_argument("output", nargs="?", type=str, help="Output file (optional), defaults to input file with appended extension")
    parser.add_argument("-c", "--csv", action="store_true", help="Save parsed data to CSV file")
    parser.add_argument("-j", "--json", action="store_true", help="Save parsed data to JSON file")
    parser.add_argument("-d", "--display", action="store_true", help="Display parsed data in console")
    parser.add_argument("-u", "--unknown", action="store_true", help="Parse input file, display unknown data to console, then exit without creating files")
    parser.add_argument("--stamp1", action="store_true", help="Prepend current UTC date/time stamp to beginning of output filenames")
    parser.add_argument("--stamp2", action="store_true", help="Append current UTC date/time stamp to end of output filename")
    parser.add_argument("--zulu", action="store_true", help="Use UTC / Zulu time for above file timestamps")
    parser.add_argument("--overwrite", action="store_true", help="Overwrite output file if it already exists")
    parser.add_argument("--preferred", action="store_true", help="Include indication of which contact fields were marked \"pref\" (duplicates value to new field)")
    parser.add_argument("--reorder", action="store_true", help="Change field order to: Organization, Name, Phones, Emails, Addresses, Dates, Relationships, (Others)")
    parser.add_argument("--address", action="store_true", help="Reformat addresses from vCard format to Postal Format")
    parser.add_argument("--phone", action="store_true", help="Reformat US phone numbers to format: NPA-NXX-XXXX")

    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f"Error: Input file '{args.input}' not found.")
        sys.exit(1)
    if not args.input.endswith(".vcf"):
        print(f"Error: Input file '{args.input}' must have a .vcf extension.")
        sys.exit(1)

    if not (args.csv or args.json or args.display or args.unknown):
        parser.error("Use --help for help; at least one of --csv, --json, --display, or --unknown must be provided.")

    format_address.enabled = args.address
    format_phone.enabled = args.phone
    parse_vcard.preferred = args.preferred

    output_file = args.output if args.output else args.input
    if output_file.endswith((".csv", ".json")):
        output_file = output_file.rsplit(".", 1)[0]

    stamp = datetime.now().strftime("%Y%m%dT%H%M%S")
    if args.zulu:
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")

    if args.stamp1:
        output_file = f"{stamp}-{output_file}"

    if args.stamp2:
        output_file = f"{output_file}-{stamp}"

    abort = False
    if args.csv and os.path.exists(output_file + ".csv") and not args.overwrite:
        print(f"Error: The file '{output_file}.csv' already exists, use --overwrite to overwrite it.")
        abort = True
    if args.json and os.path.exists(output_file + ".json") and not args.overwrite:
        print(f"Error: The file '{output_file}.json' already exists, use --overwrite to overwrite it.")
        abort = True
    if abort:
        sys.exit(1)

    custom_column_ordering.enabled = args.reorder

    data = parse_vcard(args.input, args.unknown)

    if args.unknown:
        sys.exit(1)

    if args.csv:
        save_to_csv(data, output_file + ".csv")

    if args.json:
        save_to_json(data, output_file + ".json")

    if args.display:
        print_to_console(data)

if __name__ == "__main__":
    main()
