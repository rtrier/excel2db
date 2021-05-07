package de.gdiservice.excel2db;

import java.io.File;
import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Time;
import java.sql.Timestamp;
import java.sql.Types;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.LogManager;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import de.logosib.db.ConnectionFactory;


public class Excel2DB {

    private static final de.logosib.funclib.logger.Logger logger = de.logosib.funclib.logger.LoggerFactory.getLoggerFactory().getLogger(Excel2DB.class);
	
	private String schema; 

	// private final String tablename;

	int[][] validColumns = null;
	String[][] columnNames = null; 
	Class<?>[][] columnTypes = null;

	private boolean bExcelError2Null;
	
	private boolean createSchemaIfNotExits;

    private boolean test;
	
    

    private boolean stopOnValidationError;


    private TableDescription tableDescription;
	
	Excel2DB(String schema, String importId, boolean createSchemaIfNotExits) {
		// this.fileToRead = new File(file);
		if (importId==null) {
			this.schema = schema;
		}
		else {
			this.schema = schema + "_" + importId;
		}
		this.createSchemaIfNotExits = createSchemaIfNotExits;
		
	}


	void deleteCells(Row row, int rowNr, int fromCell) {

		if (row!=null) {
			// System.out.println("\td "+rowNr+"  "+row.getLastCellNum()+"  fromCell="+fromCell);
			while (row.getLastCellNum()>fromCell) {
				int last = row.getLastCellNum()-1;
				Cell cell = row.getCell(last);
				// System.out.println("\t\t"+cell);
				if (cell!=null) {
					row.removeCell(cell);
					// System.out.println("\t\tdeleted "+last+"  "+row.getLastCellNum());
				}
				else {				
					// System.out.println("\t\t deleted break");
					break;
				}
			}
		}
	}

	String getValueAsString(Cell cell) {
		CellType cellType = cell.getCellType();
		switch (cellType) {
		case NUMERIC:
			return String.valueOf(cell.getNumericCellValue());
		case FORMULA:
			return String.valueOf(cell.getCellFormula());
		case BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue());
		default:
			return String.valueOf(cell.getStringCellValue());
		}
	}

	boolean isEmptyRow(Row row, List<ColumnDescriptor> columns) {
		int cellsWithData = 0;
		if (row!=null) {
			for (ColumnDescriptor cd : columns) {
				Cell cell = row.getCell(cd.nr);
				if (cell!=null) {
					String sValue = getValueAsString(cell);
					boolean isEmpty = sValue==null || sValue.length()==0;
					// System.out.println("  isempty "+rowNr+"  "+sValue);
					if (!isEmpty) {
						cellsWithData++;
					}
				}
			}
		}
		// System.out.println("isEmpty "+rowNr+"  "+(cellsWithData));
		return cellsWithData==0;
	}
	
	
	private boolean isEmpty(Cell cell, Class<?> clasz) {
	    try {
	        Object o = getCellValue(cell, clasz);
	        return o==null || o.toString().isBlank();
	    } 
        catch (Exception e) {
            return true;
        }
	}
	
	
	private Object getCellValue(Cell cell, Class<?> clasz) {
		// System.err.println(cell);
		
		if (cell!=null && clasz!=null) {
			// System.out.println(cell.getCellType()+"  "+bExcelError2Null);
			if (cell.getCellType()==CellType.FORMULA) {
				
				if (bExcelError2Null) {
					return null;
				}
			}
		
			if (clasz.equals(String.class)) {
			    final CellType cellType = cell.getCellType();
				if (cellType!=CellType.STRING) {
				    if (cellType==CellType.NUMERIC) {
				        double d = cell.getNumericCellValue();
				        if ((d % 1) < 0.0000001) {
				            return String.valueOf((int)d);
				        }
				        else {
				            return cell.toString();
				        }
				    } else {
				        return cell.toString();
				    }
				}
				return cell.getStringCellValue();
			} else if (clasz.equals(Double.class)) {
				return cell.getNumericCellValue();
			} else if (clasz.equals(Boolean.class)) {
				return cell.getBooleanCellValue();
			} else if (clasz.equals(java.util.Date.class)) {
				java.util.Date d = cell.getDateCellValue();
				return d==null ? null : new Timestamp(d.getTime());
			} else if (clasz.equals(java.sql.Time.class)) {
				java.util.Date d = cell.getDateCellValue();
				return d==null ? null : new Time(d.getTime());
			}
		}
		return null;
	}




	
	static public String getSQLDropTable(String tablename) {
		StringBuilder sb = new StringBuilder();
		sb.append("drop TABLE IF EXISTS ").append(tablename);
		return sb.toString();
	}
	
	static public String getSQLCreateTable(String tablename, List<ColumnDescriptor> columns) {
		StringBuilder sb = new StringBuilder();

		sb.append("CREATE TABLE ").append(tablename);
		sb.append("(");
		// sb.append("id integer NOT NULL");

		for (int i=0, count=columns.size(); i<count; i++) {
			ColumnDescriptor col = columns.get(i);
			if (i>0) {
				sb.append(",\n");
			}
			sb.append("\"").append(col.name).append("\" ");
			if (col.type==null) {
				sb.append(" varchar");
			} else if (col.type.equals(String.class)) {
				sb.append(" varchar");
			} else if (col.type.equals(Double.class)) {
				sb.append(" float");
			} else if (col.type.equals(Boolean.class)) {
				sb.append(" boolean");
			} else if (col.type.equals(java.util.Date.class)) {
				sb.append(" timestamp");
			} else if (col.type.equals(java.sql.Time.class)) {
				sb.append(" time");
			}
		}
		sb.append(")");
		return sb.toString();
	}

	static public String getSQLInsert(String tablename, List<ColumnDescriptor> columns) {
		StringBuilder sb = new StringBuilder();

		sb.append("INSERT INTO ").append(tablename);
		sb.append("(");
		// sb.append("id");

		
		for (int i=0, count=columns.size(); i<count; i++) {
			if (i>0) {
				sb.append(",");
			}
			sb.append("\"").append(columns.get(i).name).append("\"");
		}
		sb.append(") ");

		sb.append("VALUES (?");
		for (int i=1, count=columns.size(); i<count; i++) {
			sb.append(",?");
		}
		sb.append(") ");
		return sb.toString();
	}	

	boolean isHidden(Row row, int nr) {		
		return row.getZeroHeight() || (row.isFormatted() && row.getRowStyle().getHidden());
	}

	public List<SheetDescriptor> analyse(Workbook workbook) {
		System.out.println("analyse workbook");
		List<SheetDescriptor> sheetDescriptors = new ArrayList<>();
		
		for (int sheetNr=0; sheetNr<workbook.getNumberOfSheets(); sheetNr++) {
			Sheet sheet = workbook.getSheetAt(sheetNr);
			SheetDescriptor sheetDescriptor = new SheetDescriptor(sheetNr, workbook.getSheetName(sheetNr));

			
			
			int minRowNr = sheet.getFirstRowNum();
			int maxRowNr = sheet.getLastRowNum();
			
			int minColNr = Integer.MAX_VALUE;
			int maxColNr = Integer.MIN_VALUE;
			
			int firstRowNr = -1;
			
			for (int rowNr=minRowNr; rowNr<=maxRowNr; rowNr++) {
				Row row = sheet.getRow(rowNr);
				
				if (row!=null) {
					if (!isHidden(row, rowNr)) {
						if (firstRowNr<0) {
//							System.err.println("Firstrow "+rowNr);
							firstRowNr = rowNr;
							sheetDescriptor.setFirstDataRow(rowNr+1);
						}
						int firstCellNum = row.getFirstCellNum();
						if (firstCellNum>=0) {
							minColNr = Math.min(minColNr, row.getFirstCellNum());
						}
						maxColNr = Math.max(maxColNr, row.getLastCellNum());
					}
				}
			}
			
			
			
			for (int colNr=minColNr; colNr<=maxColNr; colNr++) {
				String colName = null;
				Class<?> type = null;
				// boolean hasEmptyValues = false;
				for (int rowNr=minRowNr; rowNr<=maxRowNr; rowNr++) {
					Row row = sheet.getRow(rowNr);
					if (row!=null) {
						Cell cell = row.getCell(colNr);
						if (rowNr==firstRowNr) {
							colName = getName(cell);
						}
						else {
//							if (colName!=null && colName.startsWith("Einzugsgebiet") && cell!=null)  {
//								System.out.println(rowNr+"  "+type+"  "+getType(cell)+"  \""+cell.toString()+"\"");
//							}
						    Class<?> cellType = getType(cell);
						    // hasEmptyValues = hasEmptyValues || isEmpty(cell, cellType);
							type = preferedClass(type, cellType);							
						}
					}					
				}
				if (colName!=null) {
				    if (colName.length()>0) {				
				        sheetDescriptor.addColumn(new ColumnDescriptor(colNr, colName, type));
				    } else if (type!=null) {
				        sheetDescriptor.addColumn(new ColumnDescriptor(colNr, "col_"+colNr, type));
				    }
				}
			}
			
			if (sheetDescriptor.columnDescriptors.size()>0) {
				sheetDescriptors.add(sheetDescriptor);
			}

		}
		return sheetDescriptors;
	}
	
	
	Class<?> preferedClass(Class<?> c1, Class<?> c2) {
		if (c1==null) {
			return c2;
		}
		if (c2==null) {
			return c1;
		}
		if (c1.equals(c2)) {
			return c1;
		}
		if (c1.equals(String.class) || c2.equals(String.class)) {
			return String.class;
		}
		if (c1.equals(java.util.Date.class) || c2.equals(java.util.Date.class)) {
			return java.util.Date.class;
		}
		return c1;
	}
	
	

	public void read(File fileToRead) throws Exception {
		
		FileInputStream excelFile = null;
		OPCPackage pkg = null;
		String filename = fileToRead.getName();
		String tablename = filename.substring(0, filename.lastIndexOf('.')).replaceAll("[^\\w]", "_").toLowerCase();
		
		System.out.println("Reading:\""+filename+"\"");
		
		Workbook workbook = null;
		
		List<SheetGroup> sheetGroups = new ArrayList<>();
		
		try {
			if (filename.endsWith(".xls")) {
				excelFile = new FileInputStream(fileToRead);
				workbook = new HSSFWorkbook(excelFile);
			}
			else {
				pkg = OPCPackage.open(fileToRead);
			    workbook = new XSSFWorkbook(pkg);
				// workbook = new XSSFWorkbook(excelFile);
			}
			System.out.println("Loaded:\""+filename+"\"");
			List<SheetDescriptor> sheetDescriptors = analyse(workbook);
			

			
			sheetGroups.add(new SheetGroup(sheetDescriptors.get(0)));
			for (int i=1, count=sheetDescriptors.size(); i<count; i++) {
				SheetDescriptor sd = sheetDescriptors.get(i);
				SheetGroup sg = find(sheetGroups, sd);
				if (sg==null) {
					sheetGroups.add(new SheetGroup(sd));
				}
				else {
					sg.add(sd);
				}
			}
			
            if (this.tableDescription!=null) {
                boolean isValid = validate(workbook, sheetGroups);
                if (!isValid) {
                    System.out.println("Die Valisierung war nicht erfolgreich. Bitte beachten Sie die obigen Ausgaben");
                    if (stopOnValidationError) {
                        System.out.println("Import wird abgebrochen");
                        System.exit(3);
                    }
                }
            }			
			
			for (int i=0; i<sheetGroups.size(); i++) {
			    SheetGroup sg = sheetGroups.get(i);
			    System.out.println("Gruppe " + i);
				print(sg);
			}
			if (!this.test) {
			    save(workbook, sheetGroups, tablename);
			} else {
			     // print(workbook, sheetGroups, tablename);
			}
		}
		catch (NotOfficeXmlFileException ex) {
			System.out.println("File \""+filename+"\" not a excel file.");
			ex.printStackTrace();
		}
		catch (Exception ex) {
		    System.out.println("Error processing File: [\"" + ex.getMessage() + "\"]");
		    ex.printStackTrace();
        }
		finally {
			if (workbook!=null) {
				try {
				    workbook.close();
				} catch (Exception ex) {
				    // System.out.println("Error closing workbook: [\"" + ex.getMessage() + "\"]");
                }
			}
			if (pkg!=null) {
			    try {
			        pkg.close();
                } catch (Exception ex) {
                    // System.out.println("Error closing workbook: [\"" + ex.getMessage() + "\"]");
                }
			}
		}
	}
	
	
	private boolean validate(Workbook workbook, SheetGroup sheetGroup) {
	    boolean isValid = true;
	    for (int i=0; i<sheetGroup.columnDescriptors.size(); i++) {            
            ColumnDescriptor istDescr = sheetGroup.columnDescriptors.get(i);
            ColumnDescription sollDescr = tableDescription.find(istDescr.name);
            if (sollDescr == null) {
                // isValid = false;
                // System.out.println("soll: "+sollDescr);
                System.out.println("Zielspalte für Quellspalte \""+istDescr.name+"\" wurde nicht gefunden.");
            }
        }
	    boolean checkNullConstraints = checkNullConstraint(workbook, sheetGroup);
	    return isValid && checkNullConstraints;
	}
	
	private boolean validate(Workbook workbook, List<SheetGroup> sheetGroups) {
	    boolean isValid = true;
	    System.out.println("validating mit Spalten:");	    
	    for (ColumnDescription cd : tableDescription.columns) {
	        System.out.println("\t"+cd.source_column_name+"=>"+cd.target_column_name+"  typ="+cd.data_type+" darfLeerSein="+ (cd.is_nullable? "ja":"nein"));
	    }
	    for (SheetGroup sg : sheetGroups) {
	        System.out.println("validating SheetGroup");
	        boolean isSheetgroupValid = validate(workbook, sg);
	        isValid = isValid && isSheetgroupValid;
	    }
	    if (sheetGroups.size()==1) {
	        return isValid;
	    } else {
	        System.out.println("die Strukturen der Tabellenblätter stimmen nciht überein.");
	        return false;
	    }
    }


    private void createSchemaIfNotExits(String schema) throws SQLException {
		Connection con = null;
		ResultSet rs = null;
		try {
			con = ConnectionFactory.getConnectionFactory().getConnection();
			
			rs = con.getMetaData().getSchemas(null, schema);
			if (rs.next()) {
				return;
			}
			else {
				con.createStatement().execute("create schema "+schema);
			}
		}
		finally {
			if (rs!=null) {
				try {
					rs.close();
				} catch (SQLException e) {
					e.printStackTrace();
				}
			}
			if (con!=null) {
				try {
					con.close();
				} catch (SQLException e) {
					e.printStackTrace();
				}
			}
		}
		
	}
    
    
//    private boolean checkNullConstraint(Workbook workbook, SheetGroup sg, int colNr) throws SQLException {
//        
//        
//        return false;
//    }
//    
    private boolean checkNullConstraint(Workbook workbook, SheetGroup sg) {
        boolean isValid = true;

        for (SheetDescriptor sheetDescriptor : sg.sheetDescriptors) {
            Sheet sheet = workbook.getSheetAt(sheetDescriptor.nrOfSheet);
            int maxRowNr = sheet.getLastRowNum();
            int minRowNr = sheet.getFirstRowNum();
            int firstEmptyRow = -1;

            for (int rowNr=sheetDescriptor.firstDataRow; (rowNr<=maxRowNr+1) && (firstEmptyRow<0); rowNr++) {                   
                if (rowNr!=minRowNr) {
                    Row row = sheet.getRow(rowNr);
                    if (row!=null) {
                        if (isEmptyRow(row, sg.columnDescriptors)) {
                            firstEmptyRow=rowNr;
                        }
                        else {
                            // stmt.setObject(1, sqlNr++);
                            
                            for (int colNr=0, colCount=sg.columnDescriptors.size(); colNr<colCount; colNr++) {
                                ColumnDescriptor cd = sg.columnDescriptors.get(colNr);
                                ColumnDescription cdSoll = tableDescription.find(cd.name);
                                if (cdSoll!=null && !cdSoll.is_nullable) {
                                    Cell cell = row.getCell(cd.nr);
                                    if (isEmpty(cell, cd.type)) {
                                        isValid = false;
                                        System.out.println("Arbeitsblatt \""+sheetDescriptor.sheetName+"\" kein Wert in Spalte=\""+cd.name+"\" Zeile="+(rowNr+1));
                                    }
                                }                                    
                            }
                            
                        }
                    }
                }                   
            }                   
        }                

        return isValid;
    }    

	
	private void print(Workbook workbook, List<SheetGroup> sheetGroups, String tablename) throws SQLException {
	    try {

            for (int sgNr=0, count=sheetGroups.size(); sgNr<count; sgNr++) {
                
                SheetGroup sg = sheetGroups.get(sgNr);
                sg.checkDoubleColumns();
                String cTableName = schema+".\""+ ((count==1) ? tablename : sg.sheetDescriptors.get(0).sheetName) +"\"";
                String sqlCreateTable = getSQLCreateTable(cTableName, sg.columnDescriptors);
                System.out.println(sqlCreateTable);
                String sqlInsert = getSQLInsert(cTableName, sg.columnDescriptors);
                System.out.println(sqlInsert);
                
                for (SheetDescriptor sheetDescriptor : sg.sheetDescriptors) {
                    System.out.println("writing sheet \""+sheetDescriptor.sheetName+"\" into \""+cTableName+"\"");
                    Sheet sheet = workbook.getSheetAt(sheetDescriptor.nrOfSheet);
                    int maxRowNr = sheet.getLastRowNum();
                    int minRowNr = sheet.getFirstRowNum();
                    int firstEmptyRow = -1;
                    
                    for (int rowNr=sheetDescriptor.firstDataRow; (rowNr<=maxRowNr+1) && (firstEmptyRow<0); rowNr++) {                   
                        if (rowNr!=minRowNr) {
                            Row row = sheet.getRow(rowNr);
                            if (row!=null) {
                                if (isEmptyRow(row, sg.columnDescriptors)) {
                                    firstEmptyRow=rowNr;
                                    // System.out.println("firstEmptyRow "+firstEmptyRow);                      
                                }
                                else {
                                    // stmt.setObject(1, sqlNr++);
                                    StringBuilder sb = new StringBuilder();
                                    for (int colNr=0, colCount=sg.columnDescriptors.size(); colNr<colCount; colNr++) {
                                        ColumnDescriptor cd = sg.columnDescriptors.get(colNr);
                                        Cell cell = row.getCell(cd.nr);
                                        if (colNr>0) {
                                            sb.append(", ");
                                        }
                                        if (cell!=null) {
                                            Object o;
                                            try {
                                                o = getCellValue(cell, cd.type);
                                            } 
                                            catch (Exception e) {
                                                throw new RuntimeException("sheet: \""+workbook.getSheetName(sheetDescriptor.nrOfSheet)+"\" row="+rowNr+" cell="+cd.nr+
                                                        " name=\""+cd.name+"\" cellValue=\""+cell.toString()+"\"", e);
                                            }
                                            if (o!=null) {
                                                sb.append(o);                     
                                            }
                                            else {
                                                sb.append("null"); 
                                            }
                                        }
                                        else {
                                            sb.append("null");
                                        }
                                    }
                                    System.out.println(sb);
                                }
                            
                            }
                        }                   
                    }                   
                }                
            }
	    } catch (Exception ex) {
	        ex.printStackTrace();
	    }
	}
	
	private void save(Workbook workbook, List<SheetGroup> sheetGroups, String tablename) throws SQLException {
		Connection con = null;
		try {

			con = ConnectionFactory.getConnectionFactory().getConnection();
			
			if (createSchemaIfNotExits) {
				createSchemaIfNotExits(this.schema);
			}
			
			
			for (int sgNr=0, count=sheetGroups.size(); sgNr<count; sgNr++) {
				
				SheetGroup sg = sheetGroups.get(sgNr);
				sg.checkDoubleColumns();
				
				// String cTableName = schema+"."+ ((count==1) ? tablename : tablename+"_"+sgNr);				
				String cTableName = schema+".\""+ ((count==1) ? tablename : sg.sheetDescriptors.get(0).sheetName) +"\"";	
				String sqlDropTable = getSQLDropTable(cTableName);
				// System.out.println("running \"" + sqlDropTable + "\"");
				con.createStatement().execute(sqlDropTable);
				
				
				System.out.println("writing table \""+cTableName+"\"");				
				String sqlCreateTable = getSQLCreateTable(cTableName, sg.columnDescriptors);
				System.out.println("running \"" + sqlCreateTable + "\"");
				con.createStatement().execute(sqlCreateTable);
			
				String sqlInsert = getSQLInsert(cTableName, sg.columnDescriptors);
				// System.out.println("running \"" + sqlInsert + "\"");
				PreparedStatement stmt = con.prepareStatement(sqlInsert);
				int sqlNr = 0;

				for (SheetDescriptor sheetDescriptor : sg.sheetDescriptors) {
					System.out.println("writing sheet \""+sheetDescriptor.sheetName+"\" into \""+cTableName+"\"");
					Sheet sheet = workbook.getSheetAt(sheetDescriptor.nrOfSheet);
					int maxRowNr = sheet.getLastRowNum();
					int minRowNr = sheet.getFirstRowNum();
					int firstEmptyRow = -1;
					
					for (int rowNr=sheetDescriptor.firstDataRow; (rowNr<=maxRowNr+1) && (firstEmptyRow<0); rowNr++) {					
						if (rowNr!=minRowNr) {
							Row row = sheet.getRow(rowNr);
							if (row!=null) {
								if (isEmptyRow(row, sg.columnDescriptors)) {
									firstEmptyRow=rowNr;
									// System.out.println("firstEmptyRow "+firstEmptyRow);						
								}
								else {
									// stmt.setObject(1, sqlNr++);
									for (int colNr=0, colCount=sg.columnDescriptors.size(); colNr<colCount; colNr++) {
										ColumnDescriptor cd = sg.columnDescriptors.get(colNr);
										Cell cell = row.getCell(cd.nr);
										if (cell!=null) {
											Object o;
											try {
												o = getCellValue(cell, cd.type);
											} 
											catch (Exception e) {
												throw new RuntimeException("sheet: \""+workbook.getSheetName(sheetDescriptor.nrOfSheet)+"\" row="+rowNr+" cell="+cd.nr+
														" name=\""+cd.name+"\" cellValue=\""+cell.toString()+"\"", e);
											}
											if (o!=null) {
												stmt.setObject(colNr+1, o);						
											}
											else {
												stmt.setNull(colNr+1, getSQLType(cd.type));	
											}
										}
										else {
											stmt.setNull(colNr+1, getSQLType(cd.type));
										}
									}
									stmt.addBatch();
								}
							
							}
						}					
					}
					stmt.executeBatch();
				}
				/**/
			}
		}
		finally {
			if (con!=null) {
				try {
					con.close();
				} catch (SQLException e) {
					e.printStackTrace();
				}
			}
		}
		
	}


	private int getSQLType(Class<?> clasz) {
		if (clasz==null) {
			return Types.VARCHAR;
		}
		if (clasz.equals(String.class)) {
			return Types.VARCHAR;
		}
		if (clasz.equals(Double.class)) {
			return Types.FLOAT;
		} 
		if (clasz.equals(Boolean.class)) {
			return Types.BOOLEAN;
		} 
		if (clasz.equals(java.util.Date.class)) {
			return Types.TIMESTAMP;
		}
		if (clasz.equals(java.sql.Time.class)) {
			return Types.TIME;
		}
		return -1;
	}
	
	private void print(SheetGroup sg) {
	    System.out.print("\tsheets: [");
		System.out.print(sg.sheetDescriptors.get(0).sheetName);		
		for (int i=1, count=sg.sheetDescriptors.size(); i<count; i++) {
			System.out.print(", "+sg.sheetDescriptors.get(i).sheetName);
		}
		System.out.println("]");
		System.out.println("\tColumns: ");
		List<ColumnDescriptor> columnDescriptors = sg.getColumns();
		for (ColumnDescriptor c : columnDescriptors) {
			System.out.print("\t\tNr=\""+c.nr+"\"");
			System.out.print("\tName=\""+c.name+"\"");
			System.out.println("\tType=\""+ (c.type==null ? "null" : c.type.getSimpleName()) +"\"");
		}
	}

	
	
	boolean equalsColumnNames(List<ColumnDescriptor> columnDescriptors1, List<ColumnDescriptor> columnDescriptors2) {
		if (columnDescriptors1.size()==columnDescriptors2.size()) {
			boolean result = true;
			for (int i=0, count=columnDescriptors1.size(); i<count; i++) {
				result = columnDescriptors1.get(i).nr==columnDescriptors2.get(i).nr && columnDescriptors1.get(i).name.equals(columnDescriptors2.get(i).name);
				if (!result) {
					return result;
				}
			}
			return result;
		}
		return false;
	}

	public SheetGroup find(List<SheetGroup> sheetGroups, SheetDescriptor sheetDescriptor) {
		for (SheetGroup sg : sheetGroups) {
			if (equalsColumnNames(sg.getColumns(), sheetDescriptor.columnDescriptors)) {
				return sg;
			}
		}
		return null;
	}
	
	
	private static boolean equals(List<ColumnDescriptor> columnDescriptors1, List<ColumnDescriptor> columnDescriptors2) {
		if (columnDescriptors1.size()==columnDescriptors2.size()) {
			boolean result = true;
			for (int i=0, count=columnDescriptors1.size(); i<count; i++) {
				result = columnDescriptors1.get(i).equals(columnDescriptors2.get(i));
				if (!result) {
					return result;
				}
			}
			return result;
		}
		else {
			return false;
		}
	}
	



	

	private boolean checkSheetDescriptors(List<SheetDescriptor> sheetDescriptors) {
		boolean result = true;
		SheetDescriptor firstSheetDescriptor = sheetDescriptors.get(0);
		List<ColumnDescriptor> firstColumnDescriptors = firstSheetDescriptor.columnDescriptors;
		for (int i=1, count=sheetDescriptors.size(); i<count; i++) {
			List<ColumnDescriptor> columnDescriptors = sheetDescriptors.get(i).columnDescriptors;
			result = equals(firstColumnDescriptors, columnDescriptors);
			if (!result) {
				return result;
			}
		}
		return result;
	}





	private Class[] getColumnTypes(Row row, int[] is) {
		Class[] columnTypes = new Class[is.length];
		for (int i=0, count=is.length; i<count; i++) {			
			columnTypes[i] = getType(row.getCell(is[i]));
		}
		return columnTypes;
	}

	private String[] getColumnNames(Row row, int[] is) {
		String[] columnNames = new String[is.length];
		for (int i=0, count=is.length; i<count; i++) {
			columnNames[i] = getName(row.getCell(is[i]));
		}
		return columnNames;
	}

	private String getName(Cell cell) {
		if (cell==null) {
			return null;
		}
		else {
			return cell.toString(); // .replaceAll("[^\\w]", "_");
		}
	}
	

	private Class getType(Cell cell) {
		
		if (cell==null) {
			return null;
		}
		else {
			// System.err.println(cell.toString()+"  "+cell.getCellType());
			if (CellType.NUMERIC==cell.getCellType()) {
				if (DateUtil.isCellDateFormatted(cell)) {
					java.util.Date d = cell.getDateCellValue();
					if (d==null) {
						return null;
					}
					if (d.getTime()<=24*60*1000) {
						return java.sql.Time.class;
					}
					else {
						return java.util.Date.class;
					}
				}
				return Double.class;
			};
			if (CellType.STRING==cell.getCellType()) {
				return String.class;
			};
			if (CellType.BOOLEAN==cell.getCellType()) {
				return Boolean.class;
			};
		}
		return null;
	}

	static String getNewFilename(String oldFilename) {
		int idx = oldFilename.lastIndexOf('.');
		StringBuilder sb = new StringBuilder();
		if (idx>0) {
			sb.append(oldFilename.substring(0, idx)).append("_korrigiert").append(oldFilename.substring(idx));
		}
		return sb.toString();
	}

	private int[] getValidColumns(Row row) {
		int minCellNr = row.getFirstCellNum();
		int maxCellNr = row.getLastCellNum();
		List<Integer> l = new ArrayList<Integer>();
		boolean noEmptyCells = true;
		for (int i=minCellNr; i<maxCellNr && noEmptyCells; i++) {
			Cell cell = row.getCell(i);
			String cellValue = getValueAsString(cell);
			if (cellValue!=null && cellValue.length()>0) {
				l.add(i);
			}
			else {
				noEmptyCells = false;
			}
		}
		int[] result = new int[l.size()];
		for (int i=0; i<l.size(); i++) {
			result[i] = l.get(i);
		}
		return result;
	}
	
	static class TableDescription {
	    Map<String, ColumnDescription> m = new HashMap<>();
	    List<ColumnDescription> columns = new ArrayList<>();
	    
	    void add(ColumnDescription columnDescription) {
	        m.put(columnDescription.source_column_name.toLowerCase(), columnDescription);
	        columns.add(columnDescription);
	    }
	    
	    ColumnDescription find(String sourceColumnName) {
	        return m.get(sourceColumnName.toLowerCase());
	    }
	    
	}

	static class ColumnDescriptor {
		final int nr;
		String name;
		final Class<?> type;
//		final boolean hasEmptyValues;
		
		
		public ColumnDescriptor(int nr, String colName, Class<?> type) {
            this.nr=nr;
            // this.name=(colName!=null ? colName.toLowerCase() : null);
            this.name=colName==null ? null : colName.trim();
            this.type=type;
        }
		
//		public ColumnDescriptor(int nr, String colName, Class<?> type, boolean hasEmptyValues) {
//			this.nr=nr;
//			// this.name=(colName!=null ? colName.toLowerCase() : null);
//			this.name=colName==null ? null : colName.trim();
//			this.type=type;
//			this.hasEmptyValues=hasEmptyValues;
//		}

		@Override
		public int hashCode() {
			final int prime = 31;
			int result = 1;
			result = prime * result + ((name == null) ? 0 : name.hashCode());
			result = prime * result + nr;
			result = prime * result + ((type == null) ? 0 : type.hashCode());
			return result;
		}

		@Override
		public boolean equals(Object obj) {
			if (this == obj)
				return true;
			if (obj == null)
				return false;
			if (getClass() != obj.getClass())
				return false;
			ColumnDescriptor other = (ColumnDescriptor) obj;
			if (name == null) {
				if (other.name != null)
					return false;
			} else if (!name.equals(other.name)) {				
				return false;
			}
			if (nr != other.nr)
				return false;
			if (type == null) {
				if (other.type != null)
					return false;
			} else if (!type.equals(other.type))
				return false;
			return true;
		}

        @Override
        public String toString() {
            return "ColumnDescriptor [nr=" + nr + ", name=\"" + name + "\", type=" + type + "]";
        }
		
		
	}
	
    static String normalize(String s) {
        String result = s.replaceAll("\\s", "_");
        result = result.replace("ä", "ae");
        result = result.replace("ü", "ue");
        result = result.replace("ö", "oe");
        result = result.replace("ß", "sz");
        result = result.replace("Ä", "Ae");
        result = result.replace("Ü", "Ue");
        result = result.replace("Ö", "Oe");        
        return result.replaceAll("[^\\w^\\d]", "");
    }	
	
	static class SheetDescriptor {		
		
		final String sheetName;
		final int nrOfSheet;
		int firstDataRow;
		List<ColumnDescriptor> columnDescriptors = new ArrayList<>();
		
		public SheetDescriptor(int nrOfSheet, String sheetName) {		    
			this.sheetName = normalize(sheetName);
			this.nrOfSheet = nrOfSheet;
		}

		public void setFirstDataRow(int i) {
			firstDataRow = i;
		}

		void addColumn(ColumnDescriptor columnDescriptor) {
			columnDescriptors.add(columnDescriptor);
		}	
	}
	
	static class SheetGroup {
		List<SheetDescriptor> sheetDescriptors = new ArrayList<>();
		List<ColumnDescriptor> columnDescriptors;
		
		SheetGroup(SheetDescriptor sheetDescriptor) {
			sheetDescriptors.add(sheetDescriptor);
			columnDescriptors = sheetDescriptor.columnDescriptors;
		}


		void add(SheetDescriptor sheetDescriptor) {
//			System.out.println(sheetDescriptor.sheetName);
			if (!Excel2DB.equals(sheetDescriptor.columnDescriptors, columnDescriptors)) {				
				List<ColumnDescriptor> n = unify(sheetDescriptor.columnDescriptors, columnDescriptors);				
				if (n==null) {
					throw new IllegalArgumentException("nicht die selbe Struktur");
				}
				this.columnDescriptors = n;
			}
			sheetDescriptors.add(sheetDescriptor);
		}
		
		List<ColumnDescriptor> getColumns() {
			return columnDescriptors;
		}
		
		
		public void checkDoubleColumns() {
			Map<String, List<ColumnDescriptor>> map = new HashMap<>();
			for (int colNr=1, count=columnDescriptors.size(); colNr<count; colNr++) {
				ColumnDescriptor cd = columnDescriptors.get(colNr);
				final String name = cd.name;
				for (int i=0; i<colNr; i++) {
					if (columnDescriptors.get(i).name.equals(name)) {
						List<ColumnDescriptor> l = map.get(name);
						if (l==null) {
							l = new ArrayList<>(2);
							l.add(columnDescriptors.get(i));
							map.put(name, l);
						}
						l.add(cd);
					}
				}
			}
			if (map.size()>0) {
				
				for (Map.Entry<String, List<ColumnDescriptor>> entry : map.entrySet()) {
					List<ColumnDescriptor> l = entry.getValue();
					System.out.println("Spalte \""+entry.getKey()+"\" ist "+l.size()+"x vorhanden.");
					for (int i=0; i<l.size(); i++) {
						l.get(i).name = entry.getKey() + "_" + (i+1);
					}
				}
				
			}
		}
		
		
		
		private List<ColumnDescriptor> unify(List<ColumnDescriptor> columnDescriptors1, List<ColumnDescriptor> columnDescriptors2) {
			if (columnDescriptors1.size()==columnDescriptors2.size()) {
//				System.out.println("Same size "+columnDescriptors1.size());
				List<ColumnDescriptor> columnDescriptorsNew = new ArrayList<>(columnDescriptors1.size());
				for (int i=0, count=columnDescriptors1.size(); i<count; i++) {
					String name01 = columnDescriptors1.get(i).name;
					String name02 = columnDescriptors2.get(i).name;					
					if (name01.equals(name02)) {
						Class<?> type01 = columnDescriptors1.get(i).type;
						Class<?> type02 = columnDescriptors2.get(i).type;
//						System.out.println("\tSame name \""+name01+"\"  Types:"+type01+" "+type02);
						if (type01==type02) {
							columnDescriptorsNew.add(columnDescriptors1.get(i));
						} else if (type01==null) {
							columnDescriptorsNew.add(columnDescriptors2.get(i));
						} else if (type02==null) {
							columnDescriptorsNew.add(columnDescriptors1.get(i));
						} else if (type01==String.class) {
							columnDescriptorsNew.add(columnDescriptors1.get(i));
						} else if (type02==String.class) {
							columnDescriptorsNew.add(columnDescriptors2.get(i));
						} else {
							return null;
						}
					}
					else {
						return null;
					}
				}
				return columnDescriptorsNew;
			}
			return null;
		}
		
	}


	public static void main(String[] args) {
	    LogManager.getLogManager().getLogger("").setLevel(Level.SEVERE);
	    
		try {
			ArgList argList = new ArgList(args);
			String test = argList.get("test");   
			String schema = argList.get("schema");			
			String filename = argList.get("file");
			String dirname = argList.get("dir");
			String excelError2Null = argList.get("excelError2Null");
			String importId = argList.get("importid");
			
			String importTableTypes = argList.get("importTableTypes");
			String expectedTableType = argList.get("expectedTableType");
			String sStopOnValidationError = argList.get("stopOnValidationError");
			boolean stopOnValidationError = sStopOnValidationError==null ? true : Boolean.parseBoolean(sStopOnValidationError);
			
			
			boolean bExcelError2Null = true;
			if (excelError2Null!=null) {
				bExcelError2Null = Boolean.parseBoolean(excelError2Null);				
			}
			
			String createSchemaIfNotExits = argList.get("createSchemaIfNotExits");
			boolean bCreateSchemaIfNotExits = false;
			if (createSchemaIfNotExits!=null) {
				bCreateSchemaIfNotExits = Boolean.parseBoolean(createSchemaIfNotExits);				
			}
			if (!"true".equals(test) &&  (schema==null || (filename==null && dirname==null))) {
				printVerwendung();
			}
			else if (expectedTableType!=null && importTableTypes==null) {
			    System.out.println("Error: If you specify expectedTableType you have to specify importTableTypes.");
            } else {
			 
				Excel2DB poiTest = new Excel2DB(schema, importId, bCreateSchemaIfNotExits);
				try {
				    poiTest.setValidationParam(importTableTypes, expectedTableType, stopOnValidationError);
				}
				catch (Exception ex) {
				    System.err.println("Error by getting expectedTableType: "+ex.getLocalizedMessage());
				    System.exit(1);
				}
				poiTest.setTest("true".equals(test));
				poiTest.setExcelErrors2Null(bExcelError2Null);
				// "C:\\Users\\Ralf\\ownCloud\\Austausch\\Anlage_2aa_Artdaten_MZB_STI.xlsx"
				// args = new String[] {"C:\\Users\\Ralf\\ownCloud\\2018-07-06_Bio-DB\\4. Daten\\Vorgaben und Beispieldaten\\Importdaten-Beispiel_2019_02_13\\Anlage_2bb_Artdaten_MP_STI.xlsx"};
				if (dirname!=null) {
					File dir = new File(dirname);
					if (dir.exists()) {
						String[] filenames = dir.list();
						for (String fn : filenames) {
							// if (fn.contains(".x") && !fn.equals("Anlage_3d_Begleitparameter_F.xlsx") && !fn.equals("Anlage_4c_ErgebnisÅbersicht_D.xlsx")) {
							if (fn.contains(".x")) {
								poiTest.read(new File(dir, fn));
							}
						}
					}
					else {
						System.err.println("Directory \""+dirname+"\" not found.");
					}
				}
				else {
					File file = new File(filename);
					if (file.exists()) {
						poiTest.read(file);
					}
					else {
						System.err.println("File \""+filename+"\" not found.");
					}
				}
			}
			
			ConnectionFactory.getConnectionFactory().close();
		} 
		catch (Exception ex) {
			ex.printStackTrace();
		}
	}
	

	
	private void setValidationParam(String importTableTypes, String expectedTableType, boolean stopOnValidationError) throws Exception {	     
	    this.stopOnValidationError = stopOnValidationError;
	    if (importTableTypes!=null && expectedTableType!=null) {
	        this.tableDescription = readType(importTableTypes, expectedTableType);
	    }
    }


    private TableDescription readType(String importTableTypes, String expectedTableType) throws Exception {
       Connection con = null;
        Statement stmt = null;
        ResultSet rs = null;
        try {
            con = ConnectionFactory.getConnectionFactory().getConnection();
            stmt = con.createStatement();
            
            String sql = "SELECT source_column_name, target_column_name, data_type, is_nullable\r\n"
                    + "    FROM " + importTableTypes
                    + "    where table_type = '" + expectedTableType +"'";
            rs = stmt.executeQuery(sql);
            
            final TableDescription tableDescription = new TableDescription();
            while (rs.next()) {
                final ColumnDescription colDesc = new ColumnDescription(rs);
                tableDescription.add(colDesc);
            }
            if (tableDescription.columns.size()==0) {
                throw new IllegalArgumentException("expectedTableType \""+expectedTableType+"\" not found in table \""+importTableTypes+"\".");
            }   
            return tableDescription;
        }
        finally {
            if (rs!=null) {
                try {
                    rs.close();
                } catch (SQLException e) {
                    e.printStackTrace();
                }
            }
            if (stmt!=null) {
                try {
                    stmt.close();
                } catch (SQLException e) {
                    e.printStackTrace();
                }
            }
            if (con!=null) {
                try {
                    con.close();
                } catch (SQLException e) {
                    e.printStackTrace();
                }
            }
        }
    }
    
    
    static class ColumnDescription {
        String source_column_name;
        String target_column_name;
        String data_type;
        boolean is_nullable;
        
        ColumnDescription(ResultSet rs) throws SQLException {
            int i=1;
            source_column_name= rs.getString(i++);
            target_column_name = rs.getString(i++);
            data_type = rs.getString(i++);
            is_nullable = rs.getBoolean(i++);
        }

        @Override
        public String toString() {
            return "ColumnDescription [source_column_name=" + source_column_name + ", target_column_name="
                    + target_column_name + ", data_type=" + data_type + ", is_nullable=" + is_nullable + "]";
        }
    }


    private void setTest(boolean test) {
        this.test = test;
    }


    private void setExcelErrors2Null(boolean bExcelError2Null) {
		this.bExcelError2Null = bExcelError2Null;
	}


	static void printVerwendung() {
		System.out.println("Es fehlen Parameter:\n\tschema=schema");
		System.out.println("\tfile=path2file oder dir=path (wenn dir angegeben wurde wird file ignoriert)");
		System.out.println("\t[createSchemaIfNotExits=true|false] Standard=false");
		System.out.println("\t[importid=id aus der der Schemaname gebildet wird]");
		System.out.println("\t[excelError2Null=true|false Standard=true]");
		System.out.println("to import with validation:");
		System.out.println("\t[importTableTypes=DatabaseTableWithDescription columns=(table_type, source_column_name, target_column_name, data_type, is_nullable)");
		System.out.println("\t[expectedTableType=type");
		System.out.println("\t[stopOnValidationErrors=bool Standard=true");
	}
	
	static class ArgList {
		
		Map<String, String> argMap = new HashMap<>();
		
		ArgList(String[] args) {
			if (args!=null) {
				for (int i=0; i<args.length; i++) {
					String[] sA = args[i].split("=");
					if (sA.length==2) {
						argMap.put(sA[0], sA[1]);
					}
				}
			}
		}
		
		String get(String argName) {
			return argMap.get(argName);
		}
		
	}

}
