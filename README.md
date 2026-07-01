# Solar IoT Telemetry Automation (SEMS Portal API)

An automated Python monitoring script designed to track daily generation and offline statuses for 80+ remote solar inverters. 

This tool eliminates 4+ hours of manual reporting per week byprogrammatically extracting, formatting, and color-coding status maps of weekly generation data across distributed customer installations.

## Core Features
* **API Authentication & Security:** Implements session token rotation to maintain persistent secure connections.
* **Rate-Limit Handling:** Utilizes exponential backoff to handle HTTP 429 (Rate Limit) errors gracefully during bulk data extraction.
* **Performance Optimization:** Uses local JSON caching to store device serial numbers and station IDs, significantly reducing server load and execution time.
* **Automated Data Processing:** Leverages `pandas` and `openpyxl` to safely bypass merged-cell read-only errors and dynamically write generation data into formatted Excel reports.

## Tech Stack
* Python 3
* `requests` (REST API integration)
* `pandas` & `openpyxl` (Data manipulation and Excel formatting)
