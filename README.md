# Excel2XML
Small utility that exports Excel lists to XML for easy processing with XSLT transformations.

Business loves to maintain all sorts of Excel lists. Excel sort and filter capabilities make this a conventient list store.

However when some more elaborate sorting and grouping is required and formatted reports are needed, Excel falls flat.

This utility, uses Apache POI to export an Excel file into a simplified XML format that:

 - treats the first row as column labels
 - carries row and column numbers as cell properties
 - carries the column label (if there was one) as title attribute
 
 This enables a simplified creation of XML stylesheets with XPath syntax.
 
 This isn't a sophisticated tool, but an itch I had to scratch.
 
 (C) 2017 St. Wissel, for license see LICENSE file
