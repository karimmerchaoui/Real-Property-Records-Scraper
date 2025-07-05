<div align='center'>
    
![RPR_Scraper](https://github.com/user-attachments/assets/4be3d303-e609-4449-81c7-c93c1384589e)

</div>

# Overview
The **RPR Extractor** is a powerful Python application designed to automate property research by extracting comprehensive data from RPR (Realtors Property Resource) websites.

<strong>This solves the problem</strong> of manual property research which is time-consuming and inefficient. Real estate professionals often spend hours manually searching properties one by one on RPR, copying data, and organizing information into spreadsheets.



# Who Would Benefit from This Tool:
<ul>
    <li><strong>Real Estate Professionals</strong>: Agents, brokers, and appraisers who need to gather property data efficiently</li>
    <li><strong>Property Investors</strong>: Individuals or companies analyzing multiple properties for investment opportunities</li>
    <li><strong>Real Estate Analysts</strong>: Professionals conducting market research and property valuations</li>
    <li><strong>Property Management Companies</strong>: Organizations managing large portfolios of properties</li>
    <li><strong>Real Estate Developers</strong>: Teams evaluating potential development sites</li>
    <li><strong>Mortgage Professionals</strong>: Lenders and underwriters requiring property information for loan processing</li>
</ul>

# Technologies Used
- **Python 3.x**: Core programming language
- **Selenium WebDriver**: Web automation and scraping
- **CustomTkinter**: Modern GUI framework
- **OpenPyXL**: Excel file manipulation
- **Threading & Concurrent.futures**: Multi-threaded processing for improved performance
- **Chrome WebDriver**: Browser automation engine

# Features

## Automated Data Extraction
- Automatically logs into RPR system
- Searches properties by address
- Extracts comprehensive property information
- Handles multiple properties in batch

## Comprehensive Property Data
- Estimated Property Values - Current market valuations
- House Features - Bedrooms, bathrooms, square footage, year built
- Ownership Information - Current owner details and facts
- Property Records History - Sales, financing, and transaction records
- Parcel Information - Parcel numbers and property details
- Activity Data - Recent property activity and changes

## User-Friendly Interface
- Modern dark-themed GUI
- Real-time progress tracking
- Detailed logging and status updates
- Easy file selection and output management

## Performance Optimized
- Multi-threaded processing (up to 9 concurrent threads)
- Headless browser operation for faster processing
- Efficient memory usage with optimized Chrome settings
- Automatic retry mechanisms for failed requests

## Excel Integration
- Reads input from Excel files
- Exports results to formatted Excel spreadsheets
- Maintains data structure and formatting
- Timestamped output files for organization

## Robust Error Handling
- Comprehensive exception handling
- Automatic retry for network issues
- Graceful handling of missing properties
- Detailed error logging and reporting

# Project Background
This project was originally developed for <strong>MSV Properties</strong>, a real estate company requiring efficient property data extraction and analysis capabilities.

Created by <strong>Karim Merchaoui</strong>.

# Installation

## Prerequisites
<ul>
  <li>Python 3.7 or higher</li>
  <li>Chrome browser installed on your system</li>
  <li>ChromeDriver (will be handled automatically by Selenium)</li>
</ul>

## Steps
<ul>
  <li>Clone the repository:
    <pre><code>
git clone https://github.com/yourusername/rpr-extractor.git
cd rpr-extractor
</code></pre>
  </li>
  <li>Install required dependencies:
    <pre><code>
pip install -r requirements.txt
</code></pre>
  </li>
  <li>If you don't have a requirements.txt file, install these packages manually:
    <pre><code>
pip install selenium customtkinter openpyxl urllib3 requests
</code></pre>
  </li>
</ul>

# Usage
<ol type="1">
  <li>Prepare Your Input File:
    <ul>
      <li>Create an Excel file with property addresses (one address per row)</li>
    </ul>
  </li>
  <li>Launch the application:
    <pre><code>python main.py</code></pre>
  </li>
  <li>Configure Settings:
    <ul>
      <li>Select your input Excel file</li>
      <li>Choose output folder</li>
      <li>Click "Start Processing"</li>
    </ul>
  </li>
  <li>Monitor progress in the GUI</li>
  <li>Review results in the output Excel file</li>
</ol>

# Output Data Structure
The exported Excel file contains the following comprehensive information for each property:

## General Information
| Column | Description |
|--------|-------------|
| Lowest Estimated Value | Current market valuation from RPR |
| Activity | Recent property transactions and changes |
| Parcel Number | Official property identification |
| Owner | Current property owner name |
| Mailing Address | Owner's mailing address |
| Phone Number | Owner contact number |
| Owner Occupied | Owner-occupancy status |
| Time Owned | Duration of current ownership |

## House Features
| Column | Description |
|--------|-------------|
| Property Type | Residential type (Single Family, Condo, etc.) |
| Bedrooms* | Number of bedrooms |
| Total Baths | Total bathroom count |
| Full Baths* | Full bathroom count |
| Building Area (sq ft) | Total structure square footage |
| Living Area (sq ft)* | Heated living space |
| Garage Spaces | Number of garage parking spaces |
| Year Built* | Original construction year |

## Deed Records
| Column | Description |
|--------|-------------|
| Document # | Deed document identifier |
| Contract Date | Sale contract date |
| Seller Name | Previous owner information |
| Buyer Name | Current owner at time of purchase |
| Buyer Vesting | Ownership type (Individual, Trust, LLC) |
| Recording Date | Official recording date |
| Sale Price | Last transacted price |
| Inter-family Transfer | Family transfer indicator |

## Mortgage Details
| Column | Description |
|--------|-------------|
| Loan Amount | Current mortgage balance |
| Interest Rate | Loan interest rate |
| Lender Name | Financing institution |
| Loan Type | Loan product type (Conventional, FHA, etc.) |
| Document Type | Mortgage document classification |
| Loan Term (Years) | Original loan duration |
| Due Date | Next payment due date |
| Prepayment Rider | Early payoff provisions |

## Distress Indicators
| Column | Description |
|--------|-------------|
| Document # | Distress filing identifier |
| Current Beneficiary | Lender/entity initiating action |
| Original Loan Amount | Initial loan value |
| Unpaid Balance | Outstanding debt amount |
| Case # | Court case number |
| Auction Date | Scheduled foreclosure date |
| Auction Location | Sale venue information |
| Minimum Bid Amount | Reserve price for auction |

*Full implementation includes all 80+ data points across 5 analytical categories*

![image](https://github.com/user-attachments/assets/c17aa69e-8eb1-4f07-bf35-58a39c7e7201)



# Troubleshooting
## Common Issues:
<ul>
  <li><strong>ChromeDriver Issues</strong>:
    <ul>
      <li>Ensure Chrome browser is installed and up-to-date</li>
      <li>ChromeDriver version must match your Chrome version</li>
    </ul>
  </li>
  <li><strong>Login Failures</strong>:
    <ul>
      <li>Verify RPR credentials are correct</li>
      <li>Check if RPR website structure has changed</li>
    </ul>
  </li>
  <li><strong>Property Not Found</strong>:
    <ul>
      <li>Verify address format in input Excel file</li>
      <li>Check if property exists in RPR database</li>
    </ul>
  </li>
</ul>
