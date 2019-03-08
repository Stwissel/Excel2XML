# Excel2XML
Small utility that exports Excel lists to XML for easy processing with XSLT transformations.

Business loves to maintain all sorts of Excel lists. Excel sort and filter capabilities make this a conventient list store.

However when some more elaborate sorting and grouping is required and formatted reports are needed, Excel falls flat.

This utility, uses Apache POI to export an Excel file into a simplified XML format. 
This enables a simplified creation of XML stylesheets with XPath syntax.
 
## Features 
 
<ul> 
 <li>Extract workbooks in one or separate files per worksheet</li> 
 <li>Ignore all formatting</li> 
 <li>Computed cells return their last result values, unless it is a formula error, then the formula is returned</li> 
 <li>The first line of each sheet is treated as column headers, which are extracted as columns/column elements</li> 
 <li>Each cell has a <code>column</code>, a <code>row</code> and a <code>title</code> attribute. The title reflects the value from the first row. This allows in XSLT to query the title instead of relying on the column number. Reordering, adding or removing columns won't kill your XSLT stylesheet that way</li> 
 <li>Optional empty cells can be generated with an attribute of <code>empty=&quot;true&quot;</code></li> 
 <li>Runs on Java8 completely from command line</li> 
 <li>Calling it without parameters outputs the exact syntax of options</li> 
</ul> 

## Syntax

`java -jar excel2xml.jar -i somefile.xlsx [-o somefile.xml] [-e] [-t template.xslt] [-s] [-w3,4]` 

## Parameters

<ul> 
 <li> -i the input file in xslx format</li> 
 <li> -o the output file. If missing same name as input, but extension xml</li> 
 <li> -e generate empty cells. If missing: cells without data are skipped</li> 
 <li> -s generate a single file for the whole workbook. If missing: creates one file per sheet, uses sheet name as file name</li> 
 <li> -w comma separated list of sheets to export. Starts at 0. If missing: exports all sheets. Instead of sheet number, sheet names can be used</li> 
 <li> -t optional xslt template, will run against the XML. Adjust your output file name (extension) accordingly
</ul>  
 
 This isn't a sophisticated tool, but an itch I had to scratch.
  
 (C) 2017,2019 St. Wissel, for license see LICENSE file
