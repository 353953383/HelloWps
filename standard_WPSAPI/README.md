# WPS API Documentation Scraper

This project scrapes the WPS API documentation from https://open.wps.cn/previous/docs/client/wpsLoad and stores it as text files for offline browsing.

## Setup

1. Install dependencies:
   ```
   npm install
   ```

## Usage

### Scrape Documentation

To scrape the WPS API documentation, run:
```
npm run scrape
```

This will save all documentation files to the `docs/` directory.

### View Documentation

To view the scraped documentation in a browser, run:
```
npm run view
```

Then open http://localhost:3000 in your browser.

## Structure

- `bin/scrape-wps-docs.js` - Main scraping script
- `bin/view-docs.js` - Simple server to view documentation locally
- `docs/` - Directory where scraped documentation is stored
- `package.json` - Project dependencies and scripts