package JdbcConnectivity;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EstablishConnection 
{
	//Database Driver name and Database Url
		final String dbDriver="com.mysql.jdbc.Driver";
		final String dbUrl="jdbc:mysql://localhost/distcenters";
		
		//Database credentials
		String username;
		String password;
		Connection conn=null;
		
		//File details
		String filePath="C:\\Users\\pb00001\\DcLookUp\\DistributionCenterLookUp\\data";
		String fileName="\\data.xlsx";
		String sheetName;
		
		//accept input from user
		Scanner in= new Scanner(System.in);
		public Connection connectToDatabase()
		{
			try 
			{
				Class.forName(dbDriver);
				//connect to DB
				System.out.println("Enter DB credentials");
				System.out.println("username :");
				username=in.next();
				System.out.println("Password :");
				password=in.next();
				conn=DriverManager.getConnection(dbUrl, username, password);
				System.out.println("connected to DB");
				
			}catch(ClassNotFoundException e)
			{
				System.out.println("class not found");
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				System.out.println("error connecting to DB");
			}
			return conn;
		}
		public Sheet connectToExcel() throws IOException
		{
			//Open a file
			File file =new File(filePath+fileName);
			
			//FileInputstream to read from file
			FileInputStream fis= new FileInputStream(file);
			
			//To open a Excel sheet
			
			Workbook wb=new XSSFWorkbook(fis);
			System.out.println("Enter Sheet name :");
			sheetName=in.next();
		    Sheet sh=wb.getSheet(sheetName);
		    
		    return sh;
			
		}

}
