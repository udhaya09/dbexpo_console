package test.maven.proj;

import java.io.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.sql.*;
import java.util.Scanner;

public class DBExport {

	static final String MYSQL_DRIVER = "com.mysql.jdbc.Driver";
	static final String POSTGRES_DRIVER = "org.postgresql.Driver"; 
	
	
	public static void main( String[] args ) {
		// TODO Auto-generated method stub
		getInputs();
	}

	private static void getInputs() {
		// TODO Auto-generated method stub
		String DBDriver = "";
		String userName = "";
		String password = "";
		String dbURL = "";
		String fileName  = "";
		String query  = "";
		Scanner sc = new Scanner(System.in);  
		
		System.out.println("Select DB:");
		System.out.println("1. PostgreSQL");
		System.out.println("2. MySQL");
		
		int dbType = sc.nextInt();
		
		if(dbType==1) {
			DBDriver = POSTGRES_DRIVER;			
		}
		else if (dbType==2) {
			DBDriver = MYSQL_DRIVER;	
		}
		
		System.out.println("Enter Username:");
		userName = sc.next();
		System.out.println("Enter password:");
		password = sc.next();
		if(password.equals("none")) {
			password = "";
		}
		System.out.println("Enter DB URL:");
		dbURL = sc.next();
		System.out.println("Enter target file name and location (without extension):");
		fileName = sc.next();
		fileName = fileName + ".xlsx";
		System.out.println("Enter Query:");
		Scanner in = new Scanner(System.in);
		query+= in.nextLine();
		
		getReportInExcel(DBDriver, userName, password, dbURL, fileName, query);
		
	}

	private static void getReportInExcel(String dBDriver, String userName, String password, String dbURL,
			String fileName, String query) {
		// TODO Auto-generated method stub
		try {
			//System.out.println("drivername:  " + dBDriver);
			//System.out.println("query:  " + query);
			//get results
		    Class.forName(dBDriver);
		    Connection con = DriverManager.getConnection(dbURL, userName, password);
		    Statement st = con.createStatement();
		    ResultSet rs = st.executeQuery(query);
		    
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
		    
		    //writing and saving the file
		    FileOutputStream fileOut = new FileOutputStream(fileName);
		    workbook.write(fileOut);
		    fileOut.close();
		    System.out.println("Excel report generated successfully!!");
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
