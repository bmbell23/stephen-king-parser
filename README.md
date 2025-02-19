# Stephen King Website Parser

## Version 3.0.2

A Python tool to parse and analyze Stephen King's works from his official website.

### Features
- Extracts work titles, publication dates, and types
- Creates clickable hyperlinks to original work pages
- Identifies and links to collection appearances
- Exports data to CSV with Excel-compatible hyperlinks
- Progress tracking during parsing
- Improved error handling

### Usage
```bash
stephen-king-parser --output /path/to/output
```

### Output Format
The tool generates a CSV file with the following columns:
- Read (empty column for tracking reading progress)
- Owned (empty column for tracking owned books)
- Published Date
- Title (with hyperlink to work page)
- Type
- Collection (with hyperlink to collection page, if applicable)