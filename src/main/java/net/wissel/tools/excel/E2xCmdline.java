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
		final CommandLine cmd = parser.parse(options, args);
		E2xCmdline ex = new E2xCmdline(cmd, options);
		ex.parse();
		System.out.println("Done");

	}

	private final boolean exportAllSheets;
	// Input file with extension
	private String inputFileName;
	// Output file without extension!!
	private String outputFileName;

	private final Set<Integer> sheetNumbers = new HashSet<Integer>();

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

		if (!canContinue) {
			final HelpFormatter formatter = new HelpFormatter();
			formatter.printHelp("excel2xml", options);
			System.exit(1);
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
	private void export(final XSSFSheet sheet)
			throws UnsupportedEncodingException, XMLStreamException, FactoryConfigurationError, FileNotFoundException {
		final String outputSheetName = this.outputFileName + "." + sheet.getSheetName() + OUTPUT_EXTENSION;
		final File outFile = new File(outputSheetName);
		if (outFile.exists()) {
			outFile.delete();
		}

		OutputStream outputStream = new FileOutputStream(outFile);
		XMLOutputFactory factory = XMLOutputFactory.newInstance();

		XMLStreamWriter out = factory.createXMLStreamWriter(new OutputStreamWriter(outputStream, "utf-8"));
		boolean isFirst = true;
		final Map<String, String> columns = new HashMap<String, String>();
		final String sheetName = sheet.getSheetName();
		System.out.print(sheetName);
		out.writeStartDocument();
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
		out.writeEndDocument();
		System.out.println("..");

		out.close();

	}

	/**
	 * Gets field names from column titles
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
				final String cellValue = this.getCellValue(cell, count);
				if (cellValue != null) {
					columns.put(String.valueOf(cell.getColumnIndex()), cellValue);
					out.writeStartElement("column");
					out.writeAttribute("title",cellValue);
					out.writeAttribute("col", String.valueOf(cell.getColumnIndex()));
					out.writeEndElement();
				}
				count++;
			}
			out.writeEndElement();
		} catch (XMLStreamException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * Reads the input file and exports all sheets
	 *
	 * @throws IOException
	 * @throws FactoryConfigurationError
	 * @throws XMLStreamException
	 */
	private void parse() throws IOException {
		final FileInputStream inputStream = new FileInputStream(new File(this.inputFileName));
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		int sheetCount = workbook.getNumberOfSheets();
		if (this.exportAllSheets) {
			for (int i = 0; i < sheetCount; i++) {
				final XSSFSheet sheet = workbook.getSheetAt(i);
				try {
					this.export(sheet);
				} catch (UnsupportedEncodingException | FileNotFoundException | XMLStreamException
						| FactoryConfigurationError e) {
					e.printStackTrace();
				}
			}
		} else {
			this.sheetNumbers.forEach(bigI -> {
				int i = bigI.intValue();
				if (i < sheetCount) {
					final XSSFSheet sheet = workbook.getSheetAt(i);

					try {
						this.export(sheet);
					} catch (UnsupportedEncodingException | FileNotFoundException | XMLStreamException
							| FactoryConfigurationError e) {
						e.printStackTrace();
					}
				} else {
					System.err.println("I don't have a sheet at position " + String.valueOf(i));
				}
			});
		}
		workbook.close();
		inputStream.close();
	}

	private void writeRow(final Row row, final XMLStreamWriter out, final Map<String, String> columns) {
		try {
			out.writeStartElement("row");
			final String rowNum = String.valueOf(row.getRowNum());
			out.writeAttribute("row", rowNum);
			Iterator<Cell> cellIterator = row.iterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				this.writeCell(cell, out, columns);
			}
			out.writeEndElement();
		} catch (XMLStreamException e) {
			e.printStackTrace();
		}

	}

	private void writeCell(final Cell cell, final XMLStreamWriter out, final Map<String, String> columns) {
		try {
			String cellValue = this.getCellValue(cell);
			if (cellValue != null) {
				out.writeStartElement("cell");
				String colNum = String.valueOf(cell.getColumnIndex());
				out.writeAttribute("row", String.valueOf(cell.getRowIndex()));
				out.writeAttribute("col", colNum);
				if (columns.containsKey(colNum)) {
					out.writeAttribute("title", columns.get(colNum));
				}

				if (cellValue.contains("<") || cellValue.contains(">")) {
					out.writeCData(cellValue);
				} else {
					out.writeCharacters(cellValue);
				}
				out.writeEndElement();
			}
		} catch (XMLStreamException e) {
			e.printStackTrace();
		}

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

}
