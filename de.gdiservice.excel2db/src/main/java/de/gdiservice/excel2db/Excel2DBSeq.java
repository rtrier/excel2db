/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

package de.gdiservice.excel2db;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.extractor.XSSFEventBasedExcelExtractor;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import de.gdiservice.excel2db.Excel2DB.ArgList;
import de.logosib.db.ConnectionFactory;

/**
 * A rudimentary XLSX -> CSV processor modeled on the
 * POI sample program XLS2CSVmra from the package
 * org.apache.poi.hssf.eventusermodel.examples.
 * As with the HSSF version, this tries to spot missing
 *  rows and cells, and output empty entries for them.
 * <p>
 * Data sheets are read using a SAX parser to keep the
 * memory footprint relatively small, so this should be
 * able to read enormous workbooks.  The styles table and
 * the shared-string table must be kept in memory.  The
 * standard POI styles table class is used, but a custom
 * (read-only) class is used for the shared string table
 * because the standard POI SharedStringsTable grows very
 * quickly with the number of unique strings.
 * <p>
 * For a more advanced implementation of SAX event parsing
 * of XLSX files, see {@link XSSFEventBasedExcelExtractor}
 * and {@link XSSFSheetXMLHandler}. Note that for many cases,
 * it may be possible to simply use those with a custom
 * {@link SheetContentsHandler} and no SAX code needed of
 * your own!
 */
@SuppressWarnings({"java:S106","java:S4823","java:S1192"})
public class Excel2DBSeq {




    private final OPCPackage xlsxPackage;


	private String currentSheetName;


	private RowWriter rowWriter;


	private String schema;


	private File file;


	private boolean createSchemaIfNotExits;

    /**
     * Creates a new XLSX -> CSV examples
     *
     * @param pkg        The XLSX package to process
     * @param output     The PrintStream to output the CSV to
     */
    public Excel2DBSeq(OPCPackage pkg, String schema, String filename, boolean bCreateSchemaIfNotExits) {
    	this.xlsxPackage = pkg;
    	this.schema = schema;
    	this.file = new File(filename);
    	this.createSchemaIfNotExits = bCreateSchemaIfNotExits;
	}

	/**
     * Parses and shows the content of one sheet
     * using the specified styles and shared-strings tables.
     *
     * @param styles The table of styles that may be referenced by cells in the sheet
     * @param strings The table of strings that may be referenced by cells in the sheet
     * @param sheetInputStream The stream to read the sheet-data from.

     * @exception java.io.IOException An IO exception from the parser,
     *            possibly from a byte stream or character stream
     *            supplied by the application.
     * @throws SAXException if parsing the XML data fails.
	 * @throws SQLException 
     */
    public void processSheet(
            Styles styles,
            SharedStrings strings,
            SheetContentsHandler sheetHandler,
            InputStream sheetInputStream) throws IOException, SAXException, SQLException {
        DataFormatter formatter = new DataFormatter();
        InputSource sheetSource = new InputSource(sheetInputStream);
        try {
            XMLReader sheetParser = XMLHelper.newXMLReader();
            ContentHandler handler = new XSSFSheetXMLHandler(
                  styles, null, strings, sheetHandler, formatter, false);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
            if (rowWriter!=null) {
        		rowWriter.close();
        		System.out.println(rowWriter.writtenRows);
        	}
         } catch(ParserConfigurationException e) {
            throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
         }
    }
    
    
    private void setSheetName(String sheetName) throws SQLException {
    	System.out.println("setSheetName("+sheetName+")");
    	this.currentSheetName = sheetName;    	
    	rowWriter = new RowWriter(this.schema, getTablename(sheetName), createSchemaIfNotExits);
    }

    private String getTablename(String sheetName) {
    	String fName = file.getName();
    	String sTablename = fName.substring(0, fName.lastIndexOf('.')).replaceAll("[^\\w]", "_").toLowerCase();
    	String sSheetName = sheetName.replaceAll("[^\\w]", "_").toLowerCase();
    	return "imp_"+sTablename+"_"+sSheetName;
	}

	/**
     * Initiates the processing of the XLS workbook file to CSV.
     *
     * @throws IOException If reading the data from the package fails.
     * @throws SAXException if parsing the XML data fails.
	 * @throws SQLException 
     */
    public void process() throws IOException, OpenXML4JException, SAXException, SQLException {
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(this.xlsxPackage);
        XSSFReader xssfReader = new XSSFReader(this.xlsxPackage);
        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        while (iter.hasNext()) {
            try (InputStream stream = iter.next()) {
                setSheetName(iter.getSheetName());
                processSheet(styles, strings, new SheetToDB(), stream);
            }
        }
    }
    
    /**
     * Uses the XSSF Event SAX helpers to do most of the work
     *  of parsing the Sheet XML, and outputs the contents
     *  as a (basic) CSV.
     */
    private class SheetToDB implements SheetContentsHandler {
        private boolean firstCellOfRow;
        private int currentRow = -1;
        private int currentCol = -1;
        
        private List<String> data = new ArrayList<>();
        

//        private void outputMissingRows(int number) {
//            for (int i=0; i<number; i++) {
//                for (int j=0; j<minColumns; j++) {
//                    output.append(',');
//                }
//                output.append('\n');
//            }
//        }

        @Override
        public void startRow(int rowNum) {           
            firstCellOfRow = true;
            currentRow = rowNum;
            currentCol = -1;
        }

        @Override
        public void endRow(int rowNum) {
//            // Ensure the minimum number of columns
//            for (int i=currentCol; i<minColumns; i++) {
//                output.append(',');
//            }
//            output.append('\n');
        	try {
        		rowWriter.save(data.toArray(new String[data.size()]));        		
				data.clear();
			}
        	catch (SQLException e) {
				e.printStackTrace();
			}
        }

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {

            // gracefully handle missing CellRef here in a similar way as XSSFCell does
            if(cellReference == null) {
                cellReference = new CellAddress(currentRow, currentCol).formatAsString();
            }

            // Did we miss any cells?
            int thisCol = (new CellReference(cellReference)).getCol();
            int missedCols = thisCol - currentCol - 1;
            for (int i=0; i<missedCols; i++) {
                data.add(null);
            }
            currentCol = thisCol;

            // Number or string?
            data.add(formattedValue);
//            try {
//                //noinspection ResultOfMethodCallIgnored
//                Double.parseDouble(formattedValue);
//                output.append(formattedValue);
//            } catch (Exception e) {
//                output.append('"');
//                output.append(formattedValue);
//                output.append('"');
//            }
        }
    }    

    public static void main(String[] args) throws Exception {
    	try {
			ArgList argList = new ArgList(args);
			String schema = argList.get("schema");			
			String filename = argList.get("file");
			String importId = argList.get("importid");
			boolean bExcelError2Null = true;
						
			String createSchemaIfNotExits = argList.get("createSchemaIfNotExits");
			boolean bCreateSchemaIfNotExits = false;
			if (createSchemaIfNotExits!=null) {
				bCreateSchemaIfNotExits = Boolean.parseBoolean(createSchemaIfNotExits);				
			}
			if (schema==null || (filename==null)) {
				printVerwendung();
			}
			else {

	        // The package open is instantaneous, as it should be.

			try (OPCPackage p = OPCPackage.open(filename, PackageAccess.READ)) {
		           Excel2DBSeq xlsx2csv = new Excel2DBSeq(p, schema, filename, bCreateSchemaIfNotExits);
		           xlsx2csv.process();
		    }
			}
			
			ConnectionFactory.getConnectionFactory().close();
		} 
		catch (Exception ex) {
			ex.printStackTrace();
		}


    }
    
	static void printVerwendung() {
		System.out.println("Es fehlen Parameter:\n\tschema=schema");
		System.out.println("\tfile=path2file oder dir=path (wenn dir angegeben wurde wird file ignoriert)");
		System.out.println("\t[createSchemaIfNotExits=true|false] Standard=false");
		System.out.println("\t[importid=id aus der der Schemaname gebildet wird]");
		System.out.println("\t[excelError2Null=true|false Standard=true]");
	}
}