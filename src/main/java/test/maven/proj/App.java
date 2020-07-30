package test.maven.proj;

import java.io.*;


import org.apache.poi.hssf.usermodel.HSSFWorkbook; 
import org.apache.poi.ss.usermodel.Sheet; 
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.sql.*;
/**
 * Hello world!
 *
 */
public class App 
{

	// JDBC driver name and database URL
	static final String JDBC_DRIVER = "com.mysql.jdbc.Driver";  
	static final String DB_URL = "jdbc:mysql://localhost:3306/mta_db?useTimezone=trueuseUnicode=true&useLegacyDatetimeCode=false&serverTimezone=Asia/Kolkata";

	//  Database credentials
	static final String USER = "root";
	static final String PASS = "";

	public static void main( String[] args ) throws FileNotFoundException, IOException
	{
		System.out.println( "Hello World!" );

		// Creating Workbook instances 
		Workbook wb = new HSSFWorkbook(); 

		// An output stream accepts output bytes and sends them to sink. 
		OutputStream fileOut = new FileOutputStream("Geeks.xlsx"); 

		// Creating Sheets using sheet object 
		Sheet sheet1 = wb.createSheet("Array"); 
		Sheet sheet2 = wb.createSheet("String"); 
		Sheet sheet3 = wb.createSheet("LinkedList"); 
		Sheet sheet4 = wb.createSheet("Tree"); 
		Sheet sheet5 = wb.createSheet("Dynamic Programing"); 
		Sheet sheet6 = wb.createSheet("Puzzles"); 
		System.out.println("Sheets Has been Created successfully"); 
		wb.write(fileOut); 
		getDBConnection();
		getExportExcel(JDBC_DRIVER, DB_URL, USER, PASS);

	}


	private static void getDBConnection() {
		// TODO Auto-generated method stub
		Connection conn = null;
		Statement stmt = null;
		try{
			//STEP 2: Register JDBC driver
			Class.forName("com.mysql.jdbc.Driver");

			//STEP 3: Open a connection
			System.out.println("Connecting to a selected database...");
			conn = DriverManager.getConnection(DB_URL, USER, PASS);
			System.out.println("Connected database successfully...");

			//STEP 4: Execute a query
			System.out.println("Creating statement...");
			stmt = conn.createStatement();

			String sql = "SELECT user_id, email_address, full_name FROM mta_user";
			ResultSet rs = stmt.executeQuery(sql);
			//STEP 5: Extract data from result set
			while(rs.next()){
				//Retrieve by column name
				Long id  = rs.getLong("user_id");    	         
				String email_address = rs.getString("email_address");
				String full_name = rs.getString("email_address");

				//Display values
				System.out.print("id: " + id);
				System.out.print(", email_address: " + email_address);
				System.out.println(", email_address: " + email_address);
				
			}
			rs.close();
		}catch(SQLException se){
			//Handle errors for JDBC
			se.printStackTrace();
		}catch(Exception e){
			//Handle errors for Class.forName
			e.printStackTrace();
		}finally{
			//finally block used to close resources
			try{
				if(stmt!=null)
					conn.close();
			}catch(SQLException se){
			}// do nothing
			try{
				if(conn!=null)
					conn.close();
			}catch(SQLException se){
				se.printStackTrace();
			}//end finally try
		}//end try
		System.out.println("Goodbye!");
		
	

	}
	
	private static void getExportExcel(String driver, String url, String userName, String password){
		try {
		    Class.forName(driver);
		    Connection con = DriverManager.getConnection(url, userName, password);
		    Statement st = con.createStatement();
		    ResultSet rs = st.executeQuery("SELECT * FROM mta_transaction WHERE account in (select account_id from mta_account where user = 194)");
		    System.out.println("coloumn count: " + rs.getMetaData().getColumnCount());
		    
			/*
			 * HSSFWorkbook workbook = new HSSFWorkbook(); HSSFSheet sheet =
			 * workbook.createSheet("lawix10"); HSSFRow rowhead = sheet.createRow((short)
			 * 0); rowhead.createCell((short)
			 * 0).setCellValue(rs.getMetaData().getColumnLabel(1));
			 * rowhead.createCell((short)
			 * 1).setCellValue(rs.getMetaData().getColumnLabel(2));
			 * rowhead.createCell((short)
			 * 2).setCellValue(rs.getMetaData().getColumnName(3)); int i = 1; while
			 * (rs.next()){ HSSFRow row = sheet.createRow((short) i); row.createCell((short)
			 * 0).setCellValue(rs.getString(1)); row.createCell((short)
			 * 1).setCellValue(rs.getString(2)); row.createCell((short)
			 * 2).setCellValue(rs.getString(3)); i++; } String yemi = "test.xls";
			 */
		    
		  //initiating workbook
		    XSSFWorkbook workbook = new XSSFWorkbook(); 
	        XSSFSheet sheet = workbook.createSheet("DB_EXPO_Report");
	        XSSFRow rowHead = sheet.createRow((short)0);
	        
	        //get column count
	        int columnCount = rs.getMetaData().getColumnCount();
	        
	        //create header row
	        for (int i = 0; i < columnCount; i++) {
				rowHead.createCell((short)i).setCellValue(rs.getMetaData().getColumnLabel(i+1));
			}
	        
	        //writing result records
	        int i = 1;
		    while (rs.next()){
		        XSSFRow row = sheet.createRow((short) i);
		        for (int j = 0; j < columnCount; j++) {
					row.createCell((short)j).setCellValue(rs.getString(j+1));
				}		        
		        i++;
		    }
		    String yemi = "test.xlsx";
		    FileOutputStream fileOut = new FileOutputStream(yemi);
		    workbook.write(fileOut);
		    fileOut.close();
		    
		    } catch (ClassNotFoundException e1) {
		       e1.printStackTrace();
		    } catch (SQLException e1) {
		        e1.printStackTrace();
		    } catch (FileNotFoundException e1) {
		        e1.printStackTrace();
		    } catch (IOException e1) {
		        e1.printStackTrace();
		    }
	}
}
