# My First Web Scraper in Python 🐍

Hi! I'm an IT high school senior and this is my personal project focused on automated data collection and data cleaning. 

The program automatically navigates through an online bookstore, searches for "Witcher" books, handles pagination, and extracts data across all pages. The final output is a clean, formatted Excel sheet.

## 🛠️ What the code does
* **Automated Browser Navigation:** Powered by `Playwright`, the script launches a browser, handles cookie consent banners, types keywords into the search bar, and clicks the "Next" button until the catalogue ends.
* **Robust Element Selection:** To prevent the code from breaking when the website layout changes, I designed it to locate the book title first and then look up the DOM tree (`locator("xpath=../..")`) to find the matching price.
* **Data Wrangling (Pandas):** Using regular expressions (`re.sub`), the script cleans up cluttered titles (like removing promotional labels stuck to the text) and categorizes books by their status (Bestseller, Discount, New, Damaged).
* **Custom Priority Sorting:** Data isn't dumped randomly. It follows a specific business logic priority (`Bestseller -> Discount -> Standard -> Damaged -> Out of Stock`). All unavailable books automatically rotate to the very bottom of the sheet.
* **Formatted Excel Output:** Using `openpyxl`, the code dynamically measures text length to autofit column widths so the layout looks great in LibreOffice or MS Excel. Every cell gets a clean thick border, and the header is bolded.

## 💻 Tech Stack:
* **Python 3** (Core logic)
* **Playwright** (Browser automation & rendering async JavaScript)
* **Pandas** (Data manipulation and filtering)
* **OpenPyXL** (Excel styling)
