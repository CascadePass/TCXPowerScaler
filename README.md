# TCXPowerScaler

TCXPowerScaler is a command-line utility that adjusts power values in Garmin TCX activity files. It is useful when you need to uniformly scale recorded power data after calibration changes, or equipment differences.

# Features

* Batch-processes all .tcx files in a folder.
* Preserves the original file before making changes.
* Before a file is modified, the original file is copied to: filename.tcx.original
* Supports command-line automation.
* Supports interactive entry of the scaling factor when not provided on the command line.

# Usage

## Interactive Mode

Run the executable without arguments:  TCXPowerScaler.exe

The application will:

* Use the current directory as the working folder.
* Prompt for a scaling factor.
* Process all .tcx files in that folder.

## CLI Mode

**TCXPowerScaler.exe scale:1.05**

Adjusts all TCX files in the current folder, adding 5% to each power reading.

**TCXPowerScaler.exe scale:0.97 folder:C:\Activities**

Adjusts all TCX files in C:\Activities, subtracting 3% from each power reading.

# Why This Exists

This project started with a practical problem: a cyclist needed to adjust power measurements across several years of recorded rides after switching to a device that supported a different crank length.

The proposed solution was to edit the files manually in Excel. That would have taken weeks.

Instead, I recognized that the underlying task was a straightforward data transformation and built a small utility to automate it. What would have been weeks of repetitive work became roughly an hour of development.

# Scope & Design
This is intentionally a single-file console application. It has no GUI, database, server, or external service dependencies.

The application does one job, does it locally, and doesn't require ongoing infrastructure or maintenance to remain useful.

# Project Status
TCXPowerScaler is considered complete and remains available as open-source software for anyone who needs it.

The tool has met its original requirements without needing an ongoing development roadmap.

I'm happy to address genuine issues or requests if they arise, write an issue and it will be addressed.

<img width="1536" height="1024" alt="PowerMeters" src="https://github.com/user-attachments/assets/21a95797-e1f3-468f-989d-ea57338a7dbe" />
