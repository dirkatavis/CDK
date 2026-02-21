# PFC Scrapper - Data Extraction Utility

## Purpose
Scrapes and extracts Post Final Charges (PFC) data from CDK screens for analysis and reporting.

## Entry Script
- `PFC_Scrapper.vbs` - Main data extraction script

## Input Files
- Reads from active BlueZone terminal session

## Output/Logs
- Logs written to `runtime/logs/pfc_scrapper/PFC_Scrapper.log`
- Scraped data written to `runtime/data/out/PFC_Scraped_Data.csv`

## Usage
```cmd
cscript.exe PFC_Scrapper.vbs
```

## Dependencies
- BlueZone terminal emulator with active CDK session
- `framework/PathHelper.vbs` - Path resolution
- `framework/ValidateSetup.vbs` - Setup validation
- `config/config.ini` - Configuration file

## Testing
Run app-specific tests:
```cmd
cd apps\pfc_scrapper\tests
cscript.exe test_pfc_scrapper.vbs
```

## Notes
- Extracts structured data from CDK terminal screens
- Useful for generating reports and analyzing charge patterns
- Output CSV can be processed by PowerShell utilities
