package JdbcConnectivity;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import java.sql.*;

public class ReadFromExcel
{
	EstablishConnection ec=new EstablishConnection();
	
   public void readExcel() throws SQLException 
   {
	   Connection conn=null;
	   PreparedStatement pstmt=null;
	   
	   //Connect to Database
	   
	   conn=ec.connectToDatabase();
	   conn.setAutoCommit(false);
	   String q="set foreign_key_checks=0";
	   Statement stmt=conn.createStatement();
	   stmt.executeQuery(q);
	   String query="Insert into address values(?,?,?,?,?)";
	   try {
		pstmt=conn.prepareStatement(query);
	
	   
	   //get data from Excel Sheet
	   Sheet sh=ec.connectToExcel();
	   //Rowcount
	   int rc=sh.getLastRowNum();
	   System.out.println(rc);
	   for(int k=1;k<=rc;k=k+10)
	   {
	     for(int i=k;i<k+10 && i<=rc;i++)
		   {
			   Row r=sh.getRow(i);
			   for(int j=0;j<r.getLastCellNum();j++)
			   {
				   pstmt.setString(j+1, r.getCell(j).toString());
			   }
			   pstmt.addBatch();
		   }
	       pstmt.executeBatch();
		   conn.commit();
		   System.out.println("Successfully added 10 rows");
		   
		   
	   }
	   
	   } 
	   catch (SQLException e) 
	   {
		   conn.rollback();
			System.out.println("Error in executing query");
			e.printStackTrace();
			
	   } catch (IOException e) {
		// TODO Auto-generated catch block
		   System.out.println("Error in opening file");
	}
	   q="set foreign_key_checks=1";
	   conn.createStatement();
	   stmt.executeQuery(q);
	   pstmt.close();
	   stmt.close();
	   conn.close();
	   
}
   public static void main(String[] args) throws IOException, SQLException
   {
	   ReadFromExcel obj=new ReadFromExcel();
	   obj.readExcel();
	   
   }

}
