# ORS CANADA WEB SCRAPER

This Python script scrapes product data from the ORS Canada B2B website and saves it into an Excel file and a SQLite database.

## **Features**

*   **Web Scraping:** Uses Selenium to navigate the website and extract product information, including brand, model, manufacturer part number, UPC, SKU, unit of measure (UOM), and description.
*   **Data Storage:**
    *   Stores scraped data in an Excel file (`products_data.xlsx`) with appropriate headers.
    *   Saves data to a SQLite database (`scraped_data.db`) to prevent duplicate entries.
*   **Error Handling:** Implements error handling mechanisms to manage exceptions and ensure robust execution.
*   **Progress Tracking:** Provides visual feedback on the scraping progress using the `tqdm` library.
*   **Styling:** Applies styling to the Excel output, including bold headers, auto-adjusted column widths, borders, and left-aligned descriptions.

## **Dependencies**

The script requires the following Python libraries:

*   `selenium`
*   `openpyxl`
*   `sqlite3`
*   `time`
*   `traceback`
*   `tqdm`
*   `re`

## **Usage**

1.  **Install Dependencies:** Install the required libraries using `pip install <library_name>`.
2.  **Chromedriver:** Download and place the Chromedriver executable in your system's PATH or specify its path when initializing the `Service` object.
3.  **Run the Script:** Execute the Python script. The script will:
    *   Open the target ORS Canada B2B website page.
    *   Iterate through product listings, extracting relevant information.
    *   Store the data in both an Excel file and a SQLite database.

## **How It Works**

1.  **Website Navigation:** The script navigates to the specified URL on the ORS Canada B2B website.
2.  **Brand Iteration:** It iterates through a list of brand links extracted from the page.
3.  **Product Processing:** For each brand:
    *   **Determine Total Products:** It determines the total number of products and sets up progress tracking.
    *   **Extract Product Information:**  Extracts product information (brand, model, part number, UPC, SKU, UOM, and description) from each product listing.
    *   **Data Storage:** Stores the extracted data in both an Excel file and a SQLite database, ensuring no duplicates are added.
    *   **Pagination Handling:** Handles pagination by clicking the "next page" button and continues scraping until all products for the brand have been processed.

## **Notes**

*   The script includes a `breakpoint()` statement that may be used for debugging purposes.

## **Disclaimer**

Web scraping should be done responsibly and in accordance with the website's terms of service. Ensure you have the right to scrape data from the target website before using this script.
