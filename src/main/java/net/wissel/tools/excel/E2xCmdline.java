/**
 * Command Line Tool to extract Excel Lists to XML
 * 
 * Copyright 2017 St. Wissel
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an "AS IS" BASIS, 
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. 
 * See the License for the specific language governing permissions and 
 * limitations under the License.
 */
package net.wissel.tools.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import javax.xml.stream.FactoryConfigurationError;
import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class E2xCmdline {

	private static final String OUTPUT_EXTENSION = ".xml";

	public static void main(String[] args)
			throws ParseException, IOException, XMLStreamException, FactoryConfigurationError {
		final CommandLineParser parser = new DefaultParser();
		Options options = new Options();
		options.addOption("i", "input", true, "Input xlsx File");
		options.addOption("o", "output", true, "Output XML file");
		options.addOption("w", "workbooks", true, "optional: Workbook numbers to export 0,1,2,...,n");
		options.addOption("e", "empty", false, "optional: generate tags for empty cells");
		options.addOption("s", "single", false, "optional: export all worksheets into a single output file");
		final CommandLine cmd = parser.parse(options, args);
		E2xCmdline ex = new E2xCmdline(cmd, options);
		ex.parse();
		System.out.println("Done");

	}

	private final boolean exportAllSheets;
	private final boolean exportEmptyCells;
	private final boolean exportSingleFile;

	// Input file with extension
	private String inputFileName;
	// Output file without extension!!
	private String outputFileName;

	private final Set<Integer> sheetNumbers = new HashSet<Integer>();

	/**
	 * Constructor for programatic use
	 * 
	 * @param emptyCells
	 *            Should it export empty cells
	 * @param allSheets
	 *            Should it export all sheets
	 */
	public E2xCmdline(final boolean emptyCells, final boolean allSheets) {
		this.exportAllSheets = allSheets;
		this.exportEmptyCells = emptyCells;
		this.exportSingleFile = true;
	}

	/**
	 * Constructor for command line use
	 * 
	 * @param cmd
	 *            the parameters ready parsed
	 * @param options
	 *            the expected options
	 */
	public E2xCmdline(final CommandLine cmd, final Options options) {
		boolean canContinue = true;
		if (cmd.hasOption("w")) {
			this.exportAllSheets = false;
			String[] sheetNums = cmd.getOptionValue("w").split(",");
			for (int i = 0; i < sheetNums.length; i++) {
				this.sheetNumbers.add(Integer.getInteger(sheetNums[i]));
			}
		} else {
			this.exportAllSheets = true;
		}
		if (cmd.hasOption("i")) {
			this.inputFileName = cmd.getOptionValue("i");
		} else {
			canContinue = false;
		}

		if (cmd.hasOption("o")) {
			// Strip .xml since we need the sheet number
			// before the .xml entry if we have more than one sheet
			this.outputFileName = cmd.getOptionValue("o");
			if (this.outputFileName.endsWith(OUTPUT_EXTENSION)) {
				this.outputFileName = this.outputFileName.substring(0,
						this.outputFileName.length() - OUTPUT_EXTENSION.length());
			}
		} else {
			// We add the .xml entry later anyway
			this.outputFileName = this.inputFileName;
		}

		this.exportEmptyCells = cmd.hasOption("e");
		this.exportSingleFile = cmd.hasOption("s");

		if (!canContinue) {
			final HelpFormatter formatter = new HelpFormatter();
			formatter.printHelp("excel2xml", options);
			System.exit(1);
		}
		
		if (this.exportEmptyCells) {
			System.out.println("- Generating empty cells");
		}
		if (this.exportSingleFile) {
			System.out.println("- Output to single file");
		} else {
			System.out.println("- Output to one file per sheet");
		}
		
		if (this.exportAllSheets) {
			System.out.println("- Exporting all sheets");
		} else {
			System.out.println("- Exporting selected sheets");
		}

	}

	/**
	 * Exports a single sheet to a file
	 *
	 * @param sheet
	 * @throws FactoryConfigurationError
	 * @throws XMLStreamException
	 * @throws UnsupportedEncodingException
	 * @throws FileNotFoundException
	 */
	private void export(final XSSFSheet sheet, final XMLStreamWriter out)
			throws UnsupportedEncodingException, XMLStreamException, FactoryConfigurationError, FileNotFoundException {
		boolean isFirst = true;
		final Map<String, String> columns = new HashMap<String, String>();
		final String sheetName = sheet.getSheetName();
		System.out.print(sheetName);
		out.writeStartElement("sheet");
		out.writeAttribute("name", sheetName);
		Iterator<Row> rowIterator = sheet.rowIterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			if (isFirst) {
				isFirst = false;
				this.writeFirstRow(row, out, columns);
			} else {
				this.writeRow(row, out, columns);
			}
		}
		out.writeEndElement();
		System.out.println("..");
	}

	private String getCellValue(final Cell cell) {
		return this.getCellValue(cell, -1);
	}

	private String getCellValue(final Cell cell, int count) {
		String cellValue = null;
		CellType ct = cell.getCellTypeEnum();
		switch (ct) {
		case STRING:
			cellValue = cell.getStringCellValue();
			break;
		case NUMERIC:
			cellValue = String.valueOf(cell.getNumericCellValue());
			break;
		case BOOLEAN:
			cellValue = String.valueOf(cell.getBooleanCellValue());
			break;
		case BLANK:
			if (count > -1) {
				cellValue = "BLANK" + String.valueOf(count);
			}
			break;
		case FORMULA:
			CellType cacheCellType = cell.getCachedFormulaResultTypeEnum(); {
			switch (cacheCellType) {
			case STRING:
				cellValue = cell.getStringCellValue();
				break;
			case NUMERIC:
				cellValue = String.valueOf(cell.getNumericCellValue());
				break;
			case BOOLEAN:
				cellValue = String.valueOf(cell.getBooleanCellValue());
				break;
			default:
				cellValue = cell.getCellFormula();
			}
		}
			break;
		default:
			cellValue = null;
		}
		return cellValue;
	}

	/**
	 * Create an XML Streamwriter to write into an output Stream
	 * 
	 * @param outputStream
	 *            the steam e.g. a file
	 * @return the StreamWriter
	 * @throws XMLStreamException
	 * @throws UnsupportedEncodingException
	 */
	private XMLStreamWriter getXMLWriter(final OutputStream outputStream)
			throws UnsupportedEncodingException, XMLStreamException {
		XMLOutputFactory factory = XMLOutputFactory.newInstance();
		XMLStreamWriter out = factory.createXMLStreamWriter(new OutputStreamWriter(outputStream, "utf-8"));
		return out;
	}

	private XMLStreamWriter getXMLWriter(XSSFSheet sheet)
			throws FileNotFoundException, UnsupportedEncodingException, XMLStreamException {
		final String outputSheetName = this.outputFileName + "." + sheet.getSheetName() + OUTPUT_EXTENSION;
		final File outFile = new File(outputSheetName);
		if (outFile.exists()) {
			outFile.delete();
		}
		OutputStream outputStream = new FileOutputStream(outFile);
		return this.getXMLWriter(outputStream);
	}

	/**
	 * Reads the input file and exports all sheets
	 *
	 * @throws IOException
	 * @throws FactoryConfigurationError
	 * @throws XMLStreamException
	 */
	private void parse() throws IOException, XMLStreamException {
		final InputStream inputStream = new FileInputStream(new File(this.inputFileName));
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		int sheetCount = workbook.getNumberOfSheets();
		XMLStreamWriter out = null;

		if (this.exportSingleFile) {
			final String targetFile = this.outputFileName + OUTPUT_EXTENSION;
			System.out.println("Exporting Workbook to "+targetFile);
			final File outFile = new File(targetFile);
			if (outFile.exists()) {
				outFile.delete();
			}

			out = this.getXMLWriter(new FileOutputStream(outFile));
			out.writeStartDocument();
			out.writeStartElement("workbook");
		}

		for (int i = 0; i < sheetCount; i++) {
			if (this.exportAllSheets || this.sheetNumbers.contains(String.valueOf(i))) {
				final XSSFSheet sheet = workbook.getSheetAt(i);
				if (!this.exportSingleFile) {
					out = this.getXMLWriter(sheet);
					out.writeStartDocument();
				}
				try {
					this.export(sheet, out);
				} catch (UnsupportedEncodingException | FileNotFoundException | XMLStreamException
						| FactoryConfigurationError e) {
					e.printStackTrace();
				}
				if (!this.exportSingleFile) {
					out.writeEndDocument();
					out.close();
				}
			}

		}
		// Close the XML if still open
		if (this.exportSingleFile) {
			out.writeEndElement();
			out.writeEndDocument();
		}
		if (out != null) {
			out.close();
		}
		workbook.close();
		inputStream.close();
	}

	/**
	 * Parses an inputstream containin xlsx into an outputStream containing XML
	 * 
	 * @param inputStream
	 *            the source
	 * @param outputStream
	 *            the result
	 * @throws IOException
	 * @throws XMLStreamException
	 */
	public void parse(final InputStream inputStream, final OutputStream outputStream)
			throws IOException, XMLStreamException {
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XMLStreamWriter out = this.getXMLWriter(outputStream);
		out.writeStartDocument();
		out.writeStartElement("workbook");
		int sheetCount = workbook.getNumberOfSheets();
		for (int i = 0; i < sheetCount; i++) {
			final XSSFSheet sheet = workbook.getSheetAt(i);
			try {
				this.export(sheet, out);
			} catch (UnsupportedEncodingException | FileNotFoundException | XMLStreamException
					| FactoryConfigurationError e) {
				e.printStackTrace();
			}
		}
		out.writeEndElement();
		out.writeEndDocument();
		out.close();
		workbook.close();
	}

	/**
	 * Writes out an XML cell based on coordinates and provided value
	 * 
	 * @param row
	 *            the row index of the cell
	 * @param col
	 *            the column index
	 * @param cellValue
	 *            value of the cell, can be null for an empty cell
	 * @param out
	 *            the XML output stream
	 * @param columns
	 *            the Map with column titles
	 */
	private void writeAnyCell(final int row, final int col, final String cellValue, final XMLStreamWriter out,
			final Map<String, String> columns) {
		try {
			out.writeStartElement("cell");
			String colNum = String.valueOf(col);
			out.writeAttribute("row", String.valueOf(row));
			out.writeAttribute("col", colNum);
			if (columns.containsKey(colNum)) {
				out.writeAttribute("title", columns.get(colNum));
			}
			if (cellValue != null) {
				if (cellValue.contains("<") || cellValue.contains(">")) {
					out.writeCData(cellValue);
				} else {
					out.writeCharacters(cellValue);
				}
			} else {
				out.writeAttribute("empty", "true");
			}
			out.writeEndElement();

		} catch (XMLStreamException e) {
			e.printStackTrace();
		}

	}

	/**
	 * Writes out an XML cell based on an Excel cell's actual value
	 * 
	 * @param cell
	 *            The Excel cell
	 * @param out
	 *            the output stream
	 * @param columns
	 *            the Map with column titles
	 */
	private void writeCell(final Cell cell, final XMLStreamWriter out, final Map<String, String> columns) {

		String cellValue = this.getCellValue(cell);
		int col = cell.getColumnIndex();
		int row = cell.getRowIndex();
		this.writeAnyCell(row, col, cellValue, out, columns);
	}

	/**
	 * Gets field names from column titles and writes the titles element with
	 * columns out
	 *
	 * @param row
	 *            the row to parse
	 * @param columns
	 *            the map with the values
	 */
	private void writeFirstRow(Row row, final XMLStreamWriter out, final Map<String, String> columns) {
		Iterator<Cell> cellIterator = row.iterator();
		int count = 0;
		try {
			out.writeStartElement("columns");
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				// Generate empty headers if required
				if (this.exportEmptyCells) {
					int columnIndex = cell.getColumnIndex();
					while (count < columnIndex) {
						String noLabel = "NoLabel" + String.valueOf(count);
						columns.put(String.valueOf(count), noLabel);
						out.writeStartElement("column");
						out.writeAttribute("empty", "true");
						out.writeAttribute("col", String.valueOf(count));
						out.writeAttribute("title", noLabel);
						out.writeEndElement();
						count++;
					}
				}

				final String cellValue = this.getCellValue(cell, count);
				if (cellValue != null) {
					columns.put(String.valueOf(cell.getColumnIndex()), cellValue);
					out.writeStartElement("column");
					out.writeAttribute("title", cellValue);
					out.writeAttribute("col", String.valueOf(cell.getColumnIndex()));
					out.writeEndElement();
				}
				count++;
			}
			out.writeEndElement();
		} catch (XMLStreamException e) {
			e.printStackTrace();
		}
	}

	private void writeRow(final Row row, final XMLStreamWriter out, final Map<String, String> columns) {
		try {
			int rowIndex = row.getRowNum();
			out.writeStartElement("row");
			final String rowNum = String.valueOf(rowIndex);
			out.writeAttribute("row", rowNum);
			int count = 0;
			Iterator<Cell> cellIterator = row.iterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				int columnIndex = cell.getColumnIndex();
				if (this.exportEmptyCells) {
					while (count < columnIndex) {
						this.writeAnyCell(rowIndex, count, null, out, columns);
						count++;
					}
				}
				this.writeCell(cell, out, columns);
				count++;
			}
			out.writeEndElement();
		} catch (XMLStreamException e) {
			e.printStackTrace();
		}

	}

}
