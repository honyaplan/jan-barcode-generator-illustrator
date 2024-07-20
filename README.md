# JAN Barcode Generator for Adobe Illustrator

This script generates JAN (Japanese Article Number) barcodes using Adobe Illustrator.

## Features

- Generates JAN-8 and JAN-13 barcodes
- Outputs both SVG and PNG formats
- Batch processing of multiple JAN codes
- Centering barcodes on the canvas
- Progress bar with cancellation option

## Requirements

- Adobe Illustrator
- JANCODE-nicWabun font installed

## Usage

1. Open Adobe Illustrator
2. Run the script (File > Scripts > Other Script... > select this script)
3. Enter JAN codes in the input box (one per line)
4. Click OK to start processing

The script will create a folder named "JAN" on your desktop and save the generated barcodes there.

## Notes

- The script uses batch processing to improve performance
- It includes error handling for invalid JAN codes
- The barcode size is preset to 200x90 pixels
