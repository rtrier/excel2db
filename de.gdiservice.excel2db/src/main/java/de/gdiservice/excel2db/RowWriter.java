package de.gdiservice.excel2db;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Types;
import java.util.ArrayList;
import java.util.List;

import de.gdiservice.excel2db.Excel2DB.ColumnDescriptor;
import de.logosib.db.ConnectionFactory;

public class RowWriter {
	
	private String tablename;
	private String schema;
	private boolean createSchemaIfNotExits;
	private Connection con;
	private PreparedStatement stmt; 
	private int columnCount;
	int writtenRows;
	
	RowWriter(String schema, String tablename, boolean createSchemaIfNotExits) {
		this.tablename = tablename;
		this.schema = schema;
		this.createSchemaIfNotExits = createSchemaIfNotExits;		
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
	
	private void init(String[] row) throws SQLException {
		con = ConnectionFactory.getConnectionFactory().getConnection();
		System.out.println("init");
		List<ColumnDescriptor> columnDescriptors = new ArrayList<>();
		for (int i=0, count=row.length; i<count; i++) {
			columnDescriptors.add(new ColumnDescriptor(i, row[i], String.class));
		}
		
		
		try {
			if (createSchemaIfNotExits) {
				createSchemaIfNotExits(this.schema);
			}
			String tn = this.tablename.replace("daten_registriert", "");
			
			String table = this.schema+"."+tn;
			System.out.println("init: "+table);
			// System.out.println(Excel2DB.getSQLDropTable(table));
			// System.out.println(Excel2DB.getSQLCreateTable(table, columnDescriptors));
			con.createStatement().execute(Excel2DB.getSQLDropTable(table));
			con.createStatement().execute(Excel2DB.getSQLCreateTable(table, columnDescriptors));
			
			stmt = con.prepareStatement(Excel2DB.getSQLInsert(table, columnDescriptors));
			columnCount = row.length;
		}
		catch (Exception ex) {
			if (con!=null) {
				try {
					con.close();
				} catch (SQLException e) {
					e.printStackTrace();
				}
			}
			ex.printStackTrace();
		}
		
	}

	public void save(String[] row) throws SQLException {
		if (con==null) {
			init(row);
		}
		else {
			for (int colNr=0; colNr<columnCount; colNr++) {
				Object o = (row.length>colNr) ? row[colNr] : null;
				if (o!=null) {
					stmt.setObject(colNr+1, o);						
				}
				else {
					stmt.setNull(colNr+1, Types.VARCHAR);	
				}
			}
			writtenRows++;
			stmt.addBatch();
			if (writtenRows%100 == 0) {
				stmt.executeBatch();
				System.out.println("writtenRows="+writtenRows);
			}
		}
	}
	
	public void close() throws SQLException {
		stmt.executeBatch();
		stmt.close();
		con.close();
	}
	

}
