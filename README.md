# vCard to CSV/JSON Parser

`vCardConvert.py` is a Python script to parse and convert vCard files to CSV or JSON files.

## Usage

```bash
python contacts.py input [output] [-h] [-c] [-j] [-d] [-u] [--stamp1] [--stamp2] [--zulu] [--overwrite] [--preferred] [--reorder] [--address] [--phone]
```

### Positional Arguments

| Argument        | Description                                                            |
|-----------------|------------------------------------------------------------------------|
| input           | Input file (required), the vCard file to be parsed                     |
| output          | Output file (optional), defaults to input file with appended extension |


### Options

| Option              | Description                                                                                          |
|---------------------|------------------------------------------------------------------------------------------------------|
| -h, --help          | Show this help message and exit                                                                      |
| -c, --csv           | Save parsed data to CSV file                                                                         |
| -j, --json          | Save parsed data to JSON file                                                                        |
| -d, --display       | Display parsed data in console                                                                       |
| -u, --unknown       | Parse input file, display unknown data to console, then exit without creating files                  |
|     --stamp1        | Prepend current UTC date/time stamp to beginning of output filenames                                 |
|     --stamp2        | Append current UTC date/time stamp to output filenames                                               |
|     --zulu          | Use UTC / Zulu time for above file timestamps                                                        |
|     --overwrite     | Overwrite output file if it already exists                                                           |
|     --preferred     | Include indication of which contact fields were marked "pref" (duplicates value to new field)        |
|     --reorder       | Change field order to: Organization, Name, Phones, Emails, Addresses, Dates, Relationships, (Others) |
|     --address       | Reformat addresses from vCard format to Postal Format                                                |
|     --phone         | Reformat US phone numbers to format: NPA-NXX-XXXX                                                    |


## Notes
- This script specifically targets the Apple Contacts vCard export format. Any values not present in that format's fields list will be displayed as unknown fields, use the --unknown option to check if any unknown fields exist before exporting.
- This script ignores embedded images and "sensitive content" fields.


## Examples
- Parse a vCard file and save the output as a CSV (`myfile.vcf.csv`):
```bash
python contacts.py myfile.vcf --csv
```

- Parse a vCard file and save the output as a JSON with a particular file name with suffix of the current date/time (`contacts-20260101T120000.json`), while making the addresses easier to read and standardizing the phone numbers:
```bash
python contacts.py myfile.vcf contacts --json --stamp2 --address --phone
```
