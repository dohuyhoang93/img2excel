# Image2Excel
***Currently, Image2Excel runs only on Windows. It has not been tested on Linux or other platforms.***
>***[Introduction Page](https://dohuyhoang93.github.io/img2excel/)***

## Input
1. Directory (supports local/SMB Shared) containing images.
2. Input file (txt/xlsx, supports local/SMB Shared) containing a list of product codes, one per line.
3. Search suffixes: ex: `01`, `06`.

## Functionality
The Image2Excel software will:
1. Sequentially process each product code (format: `product_code`) from the input file (txt/xlsx).  
   It searches the specified image directory (png/jpg/jpeg) for files named `product_code - suffix`. These images are copied to a new user-specified directory named `Image2Excel_Export`. (This directory must not overlap with or be inside the scanned image directory; otherwise, Image2Excel will not run and will display a warning.)  
2. Create a new Excel file named `output_date_time.xlsx` (where `date_time` is the current date and time).  
   For each product code, insert the `product_code` and its corresponding images (`product_code - suffix_01`, `product_code - suffix_06`) into cells on the same row, starting from cell A2. Save the Excel file.  
   Repeat for all product codes in the input file.

## Output
1. A directory containing images that match the product codes from the input file.
2. An Excel file containing columns for product code, image `01`, and image `06` for all product codes in the input file.

Example Excel output layout:
|A|B|C|D|E|
|---|---|---|---|---|
||||||
|abc0123|abc012-01.jpg|abc012-06.jpg|
|abc456|abc456-01.jpg|abc456-06.jpg|

## Other
**Theme:**
 * Toggle between dark and light themes.<br>

**JSON**
 * User settings are saved in a JSON file located at: `%APPDATA%\Image2Excel\settings.json`

## Planned Updates
1. Synchronize Menu bar color with the selected theme.
2. Improve the positioning of the filter canvas for better aesthetics.
3. Add a user guide to the Help menu.