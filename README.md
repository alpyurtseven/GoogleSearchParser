# GoogleSearchParser

GoogleSearchParser is a simple application that utilizes the HTMLAgilityPack NuGet package to search for company names listed in an Excel file on Google and retrieve information such as the company name, phone number (if available), opening hours, and website address from Google Business Data. It then saves this information to an Excel file.

## Overview

The application provides a convenient way to extract essential business information from Google search results for multiple companies listed in an Excel file. This can be useful for various purposes such as lead generation, market analysis, or data collection.

## How it Works

1. The application prompts the user to provide the file path of an Excel file containing a list of company names.
2. It reads the Excel file and extracts the list of company names.
3. For each company name in the list, it performs a Google search.
4. It parses the Google search results to extract relevant business data using the HTMLAgilityPack library.
5. Information such as company name, phone number, opening hours, and website address is collected from Google Business Data.
6. The extracted data is saved to an Excel file for further analysis or reference.

## Installation

1. Clone this repository to your local machine.
2. Open the solution in Visual Studio or your preferred IDE.
3. Build the solution to restore NuGet packages.
4. Run the application.

## Usage

1. Prepare an Excel file (.xlsx) containing a list of company names in one of the columns.
2. Run the application.
3. When prompted, enter the file path of the Excel file containing the list of company names.
4. Wait for the application to process the search results and extract the business information.
5. Once the process is complete, the extracted data will be saved to excel file that you provided.

