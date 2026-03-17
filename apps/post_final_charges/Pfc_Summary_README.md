# Pfc_Summary - PFC Status Scraper

## Purpose
Loops through a configured range of PFC sequence numbers, scrapes the RO number
and status from each screen, and writes the results to a CSV file. Designed to
quickly build a snapshot of all RO statuses across a sequence range without
performing any write operations on the CDK system.

## Entry Script
- `Pfc_Summary.vbs` - Main scraper script

## Input
- Active BlueZone terminal session positioned at the PFC screen

## Output / Logs
All paths are defined in `config/config.ini` under `[Pfc_Summary]`:

| Setting               | Description                                 |
|-----------------------|---------------------------------------------|
| `OutputCSV`           | CSV file written with one row per RO        |
| `Log`                 | Execution log                               |

CSV format:
```
RO_Number,Status
860025,READY TO POST
872511,OPENED
```

## Configuration (`config/config.ini`)

```ini
[Pfc_Summary]
OutputCSV=apps\post_final_charges\Pfc_Summary_output.csv
Log=apps\post_final_charges\Pfc_Summary.log
StartSequenceNumber=1
EndSequenceNumber=100
StepDelayMs=1000
```

| Key                   | Description                                               |
|-----------------------|-----------------------------------------------------------|
| `StartSequenceNumber` | First sequence number to process                          |
| `EndSequenceNumber`   | Last sequence number to process (inclusive)               |
| `StepDelayMs`         | Milliseconds to wait between screen transitions (default 1000) |

## Usage

Run from the BlueZone Script Host or cscript:

```cmd
cscript.exe apps\post_final_charges\Pfc_Summary.vbs
```

## How It Works

For each sequence in the configured range:
1. Waits for the `COMMAND:` prompt
2. Sends the sequence number + Enter
3. Waits for the RO detail screen (`RO STATUS:` marker)
4. Scrapes RO number and status
5. Writes one CSV row
6. Sends `E` + Enter to return to the command prompt
7. Stops early if `DOES NOT EXIST` is encountered (end of data)

## Dependencies
- BlueZone terminal emulator with active CDK session
- `framework/PathHelper.vbs` - Path resolution
- `config/config.ini` - Configuration

## Notes
- Read-only: makes no changes to the CDK system
- Safe to run repeatedly; output CSV is overwritten each run
- `StepDelayMs` can be increased if the terminal is slow to respond
