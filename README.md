# My First Web Scraper in Python 🐍 (Now with Image Downloads!)

Hi! I'm an IT high school senior and this is my first project focused on automated data collection, data cleaning, and Excel reporting. 

The program automatically navigates through an online bookstore, searches for "Witcher" books, handles pagination, and extracts data across all pages. The final output is a clean, formatted Excel sheet that now includes actual book covers embedded directly inside the cells!

## 🛠️ What the code does
* **Automated Browser Navigation:** Powered by `Playwright`, the script launches a browser, handles cookie consent banners, types keywords into the search bar, and clicks the "Next" button until the catalogue ends.
* **Robust Element Selection:** To prevent the code from breaking when the website layout changes, I designed it to locate the book title first and then look up the DOM tree (`locator("xpath=../..")`) to find the matching price.
* **Smart Image Scraping & Quality Fix:** The script bypasses the website's lazy loading by forcing each book cover into view (`scroll_into_view_if_needed`). It extracts high-resolution images from the `srcset` attribute, cleans forbidden Windows characters from filenames (`re.sub`), and downloads them locally into a folder.
* **Data Wrangling (Pandas):** Using regular expressions, the script cleans up cluttered titles (like removing promotional labels stuck to the text) and categorizes books by their status (Bestseller, Discount, New, Damaged).
* **Custom Priority Sorting:** Data follows a specific business logic priority (`Bestseller -> Discount -> Standard -> Damaged -> Out of Stock`). All unavailable books automatically rotate to the very bottom of the sheet.
* **Embedded Image Excel Output:** Using `openpyxl`, the code dynamically measures text length to autofit column widths. It automatically adjusts row heights to fit visual elements and embeds the downloaded images directly into Column E while cleanly hiding cell borders for that specific column so the covers look amazing.

## 💻 Tech Stack:
* **Python 3** (Core logic)
* **Playwright** (Browser automation & rendering async JavaScript)
* **Pandas** (Data manipulation and filtering)
* **OpenPyXL** (Excel styling, layout manipulation & drawing images)
