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
	static final String JDBC_DRIVER = "org.postgresql.Driver";  
	static final String DB_URL = "jdbc:postgresql://localhost:54321/wem_prod_db";

	//  Database credentials
	static final String USER = "nlproddb";
	static final String PASS = "0xtoWNVTmjFa5IJD8CL8";

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
		//getDBConnection();
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

			String sql = "SELECT DISTINCT C.CUSTOMER_ID, TO_CHAR(C.DATE_CREATED,'MM/DD/YYYY') AS DATE_CREATED,\r\n" + 
					"case when LAST_LOGIN_DATE='' then null else\r\n" + 
					"TO_CHAR(TO_DATE(LAST_LOGIN_DATE, 'YYYY/MM/DD'),'MM/DD/YYYY') end AS LAST_LOGIN_DATE,\r\n" + 
					"C.IS_REGISTERED, C.DEACTIVATED, C.USER_NAME, C.EMAIL_ADDRESS, C.FIRST_NAME, C.LAST_NAME,\r\n" + 
					"CAA.VALUE AS JOB_TITLE, CAB.VALUE AS PHONE_XPRESS, CAC.VALUE AS NEED_BUSINESS_INFO, \r\n" + 
					"CAD.VALUE AS NEED_NEWSLETTER, CC.COMPANY_ID AS COMPANY_SALESFORCE_ID, CAE.VALUE AS CONTACT_SF_ID,\r\n" + 
					"CAF.VALUE AS COMPANY_NAME_XPRESS, CAG.VALUE AS COMPANY_TYPE_XPRESS, CAH.VALUE AS STREET_XPRESS, CAI.VALUE AS STATE, \r\n" + 
					"CAJ.VALUE AS CITY, CAK.VALUE AS ZIPCODE, CAL.VALUE AS COUNTRY, \r\n" + 
					"CC.SALES_REP_ID, CC.REFERING_DOMAIN, C.RECEIVE_EMAIL, CC.SITE_NAME, \r\n" + 
					"CC.IS_SALESREP,\r\n" + 
					"CAM.VALUE AS SUBSCRIPTION_TYPE, \r\n" + 
					"CASE WHEN SUBSTR(CAN.VALUE,1,2) != '20' THEN CAM.VALUE \r\n" + 
					"ELSE TO_CHAR(TO_DATE(CAN.VALUE, 'YYYY/MM/DD'),'MM/DD/YYYY') END AS SUBSCRIPTION_EXPIRES,\r\n" + 
					"CAO.VALUE AS SUBSCRIPTION_IMP,\r\n" + 
					"CASE WHEN SUBSTR(CAP.VALUE,1,2) != '20' THEN CAO.VALUE \r\n" + 
					"ELSE TO_CHAR(TO_DATE(CAP.VALUE, 'YYYY/MM/DD'),'MM/DD/YYYY') END as SUBSCRIPTION_IMP_EXPIRY,\r\n" + 
					"C.CREATED_BY,  TO_CHAR(C.DATE_UPDATED,'MM/DD/YYYY') AS DATE_UPDATED, C.UPDATED_BY, \r\n" + 
					"STRING_AGG(CASE WHEN CPX.CUSTOMER_PERMISSION_ID=-1 THEN 'VIEW CLIPS'\r\n" + 
					"		WHEN CPX.CUSTOMER_PERMISSION_ID=-2 THEN 'VIEW PREVIEW'\r\n" + 
					"        WHEN CPX.CUSTOMER_PERMISSION_ID=-3 THEN 'VIEW COMPILATIONS'\r\n" + 
					"        WHEN CPX.CUSTOMER_PERMISSION_ID=-4 THEN 'VIEW TEXT RECORDS'\r\n" + 
					"        WHEN CPX.CUSTOMER_PERMISSION_ID=-5 THEN 'VIEW TRANSCRIPTS'\r\n" + 
					"        WHEN CPX.CUSTOMER_PERMISSION_ID=-6 THEN 'VIEW PHOTO'\r\n" + 
					"        WHEN CPX.CUSTOMER_PERMISSION_ID=-7 THEN 'VIEW ILLUSTRATION'\r\n" + 
					"        WHEN CPX.CUSTOMER_PERMISSION_ID=-8 THEN 'IMPERSONATE PERMISSION'\r\n" + 
					"        WHEN CPX.CUSTOMER_PERMISSION_ID=-9 THEN 'MANAGE QUOTE'\r\n" + 
					"        WHEN CPX.CUSTOMER_PERMISSION_ID=-10 THEN 'ADMINISTER CART'\r\n" + 
					"        ELSE 'NO PERMISSIONS' END\r\n" + 
					"       ,',')  AS PERMISSIONS\r\n" + 
					"FROM wem_mgmt_schema.BLC_CUSTOMER C \r\n" + 
					"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAA ON C.CUSTOMER_ID = CAA.CUSTOMER_ID AND CAA.NAME = 'job_title'\r\n" + 
					"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAB ON C.CUSTOMER_ID = CAB.CUSTOMER_ID AND CAB.NAME = 'phoneno'\r\n" + 
					"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAC ON C.CUSTOMER_ID = CAC.CUSTOMER_ID AND CAC.NAME LIKE 'NeedB%'\r\n" + 
					"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAD ON C.CUSTOMER_ID = CAD.CUSTOMER_ID AND CAD.NAME LIKE 'NeedNe%'\r\n" + 
					"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAE ON C.CUSTOMER_ID = CAE.CUSTOMER_ID AND CAE.NAME = 'contactsfID'\r\n" + 
					"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAF ON C.CUSTOMER_ID = CAF.CUSTOMER_ID AND CAF.NAME = 'companyName'\r\n" + 
					"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAG ON C.CUSTOMER_ID = CAG.CUSTOMER_ID AND CAG.NAME = 'companyType'\r\n" + 
					"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAH ON C.CUSTOMER_ID = CAH.CUSTOMER_ID AND CAH.NAME = 'address'\r\n" + 
					"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAI ON C.CUSTOMER_ID = CAI.CUSTOMER_ID AND CAI.NAME = 'state'\r\n" + 
					"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAJ ON C.CUSTOMER_ID = CAJ.CUSTOMER_ID AND CAJ.NAME = 'city'\r\n" + 
					"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAK ON C.CUSTOMER_ID = CAK.CUSTOMER_ID AND CAK.NAME = 'zipcode'\r\n" + 
					"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAL ON C.CUSTOMER_ID = CAL.CUSTOMER_ID AND CAL.NAME = 'country'\r\n" + 
					"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAM ON C.CUSTOMER_ID = CAM.CUSTOMER_ID AND CAM.NAME = 'subscription_type'\r\n" + 
					"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAN ON C.CUSTOMER_ID = CAN.CUSTOMER_ID AND CAN.NAME = 'subscription_expiry'\r\n" + 
					"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAO ON C.CUSTOMER_ID = CAO.CUSTOMER_ID AND CAO.NAME = 'subscription_imp'\r\n" + 
					"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAP ON C.CUSTOMER_ID = CAP.CUSTOMER_ID AND CAP.NAME = 'subscription_impexpiry'\r\n" + 
					"LEFT JOIN wem_mgmt_schema.BLC_CUSTOM_CUSTOMER CC ON C.CUSTOMER_ID = CC.CUSTOMER_ID\r\n" + 
					"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_PERMISSION_XREF CPX ON C.CUSTOMER_ID = CPX.CUSTOMER_ID\r\n" + 
					"WHERE CC.SITE_NAME LIKE 'XPRESS' AND TO_CHAR(C.DATE_CREATED,'MM/YYYY') = '01/2020' \r\n" + 
					"GROUP BY C.CUSTOMER_ID, C.DATE_CREATED, CC.LAST_LOGIN_DATE, C.IS_REGISTERED, C.DEACTIVATED, C.USER_NAME,\r\n" + 
					"C.EMAIL_ADDRESS, C.FIRST_NAME, C.LAST_NAME, CAA.VALUE, CAB.VALUE, CAC.VALUE,\r\n" + 
					"CAD.VALUE, CC.COMPANY_ID, CAE.VALUE, CAF.VALUE, CAG.VALUE, CAH.VALUE, CAI.VALUE, CAJ.VALUE,\r\n" + 
					"CAK.VALUE, CAL.VALUE, CC.IS_SALESREP,\r\n" + 
					"CC.SALES_REP_ID, CC.REFERING_DOMAIN, C.RECEIVE_EMAIL, CC.SITE_NAME, \r\n" + 
					"CAM.VALUE, CAN.VALUE,  CAO.VALUE, CAP.VALUE,\r\n" + 
					"C.CREATED_BY, C.DATE_UPDATED, C.UPDATED_BY\r\n" + 
					"ORDER BY C.customer_id\r\n" + 
					"";
			ResultSet rs = stmt.executeQuery(sql);
			//STEP 5: Extract data from result set
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
			    FileOutputStream fileOut = new FileOutputStream("text.xlsx");
			    workbook.write(fileOut);
			    fileOut.close();
			    System.out.println("Excel report generated successfully!!");
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
		    ResultSet rs = st.executeQuery("SELECT DISTINCT C.CUSTOMER_ID, TO_CHAR(C.DATE_CREATED,'MM/DD/YYYY') AS DATE_CREATED,\r\n" + 
		    		"case when LAST_LOGIN_DATE='' then null else\r\n" + 
		    		"TO_CHAR(TO_DATE(LAST_LOGIN_DATE, 'YYYY/MM/DD'),'MM/DD/YYYY') end AS LAST_LOGIN_DATE,\r\n" + 
		    		"C.IS_REGISTERED, C.DEACTIVATED, C.USER_NAME, C.EMAIL_ADDRESS, C.FIRST_NAME, C.LAST_NAME,\r\n" + 
		    		"CAA.VALUE AS JOB_TITLE, CAB.VALUE AS PHONE_XPRESS, CAC.VALUE AS NEED_BUSINESS_INFO, \r\n" + 
		    		"CAD.VALUE AS NEED_NEWSLETTER, CC.COMPANY_ID AS COMPANY_SALESFORCE_ID, CAE.VALUE AS CONTACT_SF_ID,\r\n" + 
		    		"CAF.VALUE AS COMPANY_NAME_XPRESS, CAG.VALUE AS COMPANY_TYPE_XPRESS, CAH.VALUE AS STREET_XPRESS, CAI.VALUE AS STATE, \r\n" + 
		    		"CAJ.VALUE AS CITY, CAK.VALUE AS ZIPCODE, CAL.VALUE AS COUNTRY, \r\n" + 
		    		"CC.SALES_REP_ID, CC.REFERING_DOMAIN, C.RECEIVE_EMAIL, CC.SITE_NAME, \r\n" + 
		    		"CC.IS_SALESREP,\r\n" + 
		    		"CAM.VALUE AS SUBSCRIPTION_TYPE, \r\n" + 
		    		"CASE WHEN SUBSTR(CAN.VALUE,1,2) != '20' THEN CAM.VALUE \r\n" + 
		    		"ELSE TO_CHAR(TO_DATE(CAN.VALUE, 'YYYY/MM/DD'),'MM/DD/YYYY') END AS SUBSCRIPTION_EXPIRES,\r\n" + 
		    		"CAO.VALUE AS SUBSCRIPTION_IMP,\r\n" + 
		    		"CASE WHEN SUBSTR(CAP.VALUE,1,2) != '20' THEN CAO.VALUE \r\n" + 
		    		"ELSE TO_CHAR(TO_DATE(CAP.VALUE, 'YYYY/MM/DD'),'MM/DD/YYYY') END as SUBSCRIPTION_IMP_EXPIRY,\r\n" + 
		    		"C.CREATED_BY,  TO_CHAR(C.DATE_UPDATED,'MM/DD/YYYY') AS DATE_UPDATED, C.UPDATED_BY, \r\n" + 
		    		"STRING_AGG(CASE WHEN CPX.CUSTOMER_PERMISSION_ID=-1 THEN 'VIEW CLIPS'\r\n" + 
		    		"		WHEN CPX.CUSTOMER_PERMISSION_ID=-2 THEN 'VIEW PREVIEW'\r\n" + 
		    		"        WHEN CPX.CUSTOMER_PERMISSION_ID=-3 THEN 'VIEW COMPILATIONS'\r\n" + 
		    		"        WHEN CPX.CUSTOMER_PERMISSION_ID=-4 THEN 'VIEW TEXT RECORDS'\r\n" + 
		    		"        WHEN CPX.CUSTOMER_PERMISSION_ID=-5 THEN 'VIEW TRANSCRIPTS'\r\n" + 
		    		"        WHEN CPX.CUSTOMER_PERMISSION_ID=-6 THEN 'VIEW PHOTO'\r\n" + 
		    		"        WHEN CPX.CUSTOMER_PERMISSION_ID=-7 THEN 'VIEW ILLUSTRATION'\r\n" + 
		    		"        WHEN CPX.CUSTOMER_PERMISSION_ID=-8 THEN 'IMPERSONATE PERMISSION'\r\n" + 
		    		"        WHEN CPX.CUSTOMER_PERMISSION_ID=-9 THEN 'MANAGE QUOTE'\r\n" + 
		    		"        WHEN CPX.CUSTOMER_PERMISSION_ID=-10 THEN 'ADMINISTER CART'\r\n" + 
		    		"        ELSE 'NO PERMISSIONS' END\r\n" + 
		    		"       ,',')  AS PERMISSIONS\r\n" + 
		    		"FROM wem_mgmt_schema.BLC_CUSTOMER C \r\n" + 
		    		"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAA ON C.CUSTOMER_ID = CAA.CUSTOMER_ID AND CAA.NAME = 'job_title'\r\n" + 
		    		"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAB ON C.CUSTOMER_ID = CAB.CUSTOMER_ID AND CAB.NAME = 'phoneno'\r\n" + 
		    		"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAC ON C.CUSTOMER_ID = CAC.CUSTOMER_ID AND CAC.NAME LIKE 'NeedB%'\r\n" + 
		    		"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAD ON C.CUSTOMER_ID = CAD.CUSTOMER_ID AND CAD.NAME LIKE 'NeedNe%'\r\n" + 
		    		"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAE ON C.CUSTOMER_ID = CAE.CUSTOMER_ID AND CAE.NAME = 'contactsfID'\r\n" + 
		    		"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAF ON C.CUSTOMER_ID = CAF.CUSTOMER_ID AND CAF.NAME = 'companyName'\r\n" + 
		    		"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAG ON C.CUSTOMER_ID = CAG.CUSTOMER_ID AND CAG.NAME = 'companyType'\r\n" + 
		    		"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAH ON C.CUSTOMER_ID = CAH.CUSTOMER_ID AND CAH.NAME = 'address'\r\n" + 
		    		"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAI ON C.CUSTOMER_ID = CAI.CUSTOMER_ID AND CAI.NAME = 'state'\r\n" + 
		    		"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAJ ON C.CUSTOMER_ID = CAJ.CUSTOMER_ID AND CAJ.NAME = 'city'\r\n" + 
		    		"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAK ON C.CUSTOMER_ID = CAK.CUSTOMER_ID AND CAK.NAME = 'zipcode'\r\n" + 
		    		"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAL ON C.CUSTOMER_ID = CAL.CUSTOMER_ID AND CAL.NAME = 'country'\r\n" + 
		    		"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAM ON C.CUSTOMER_ID = CAM.CUSTOMER_ID AND CAM.NAME = 'subscription_type'\r\n" + 
		    		"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAN ON C.CUSTOMER_ID = CAN.CUSTOMER_ID AND CAN.NAME = 'subscription_expiry'\r\n" + 
		    		"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAO ON C.CUSTOMER_ID = CAO.CUSTOMER_ID AND CAO.NAME = 'subscription_imp'\r\n" + 
		    		"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_ATTRIBUTE CAP ON C.CUSTOMER_ID = CAP.CUSTOMER_ID AND CAP.NAME = 'subscription_impexpiry'\r\n" + 
		    		"LEFT JOIN wem_mgmt_schema.BLC_CUSTOM_CUSTOMER CC ON C.CUSTOMER_ID = CC.CUSTOMER_ID\r\n" + 
		    		"LEFT JOIN wem_mgmt_schema.BLC_CUSTOMER_PERMISSION_XREF CPX ON C.CUSTOMER_ID = CPX.CUSTOMER_ID\r\n" + 
		    		"WHERE CC.SITE_NAME LIKE 'XPRESS' AND TO_CHAR(C.DATE_CREATED,'MM/YYYY') = '01/2020' \r\n" + 
		    		"GROUP BY C.CUSTOMER_ID, C.DATE_CREATED, CC.LAST_LOGIN_DATE, C.IS_REGISTERED, C.DEACTIVATED, C.USER_NAME,\r\n" + 
		    		"C.EMAIL_ADDRESS, C.FIRST_NAME, C.LAST_NAME, CAA.VALUE, CAB.VALUE, CAC.VALUE,\r\n" + 
		    		"CAD.VALUE, CC.COMPANY_ID, CAE.VALUE, CAF.VALUE, CAG.VALUE, CAH.VALUE, CAI.VALUE, CAJ.VALUE,\r\n" + 
		    		"CAK.VALUE, CAL.VALUE, CC.IS_SALESREP,\r\n" + 
		    		"CC.SALES_REP_ID, CC.REFERING_DOMAIN, C.RECEIVE_EMAIL, CC.SITE_NAME, \r\n" + 
		    		"CAM.VALUE, CAN.VALUE,  CAO.VALUE, CAP.VALUE,\r\n" + 
		    		"C.CREATED_BY, C.DATE_UPDATED, C.UPDATED_BY\r\n" + 
		    		"ORDER BY C.customer_id\r\n" + 
		    		"");
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
