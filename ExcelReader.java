
package com;

import java.sql.Statement;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

/**
 *
 * @author Parth
 */
public class ExcelReader {

    /**
     * Reading and writing Excel files using JDBC
     */
    public static void main(String[] args) {
        Connection con;
        try {
            con = DriverManager.getConnection(
                    "jdbc:odbc:" +
                    "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};" +
                    "DBQ=C:\\forestfires.xlsx;" +
                    "ReadOnly=0;");
            Statement s = con.createStatement();
            ResultSet rs = s.executeQuery("SELECT * FROM [Sheet1$]");
            rs.next();
            System.out.println(rs.getString("temp"));
            // leave rs open
            
            //Updating, original value of temp=8.2
            PreparedStatement ps = con.prepareStatement(
                    "UPDATE [Sheet1$] SET temp=? WHERE X=?");
            ps.setString(1, "8.7");
            ps.setInt(2, 7);
            ps.executeUpdate();
            ps.close();
            System.out.println("temp changed to 8.7");

            // do another read from rs
            rs.next();
            System.out.println(rs.getString("temp"));

            rs.close();
            s.close();
            con.close();
        } catch (SQLException ex) {
            ex.printStackTrace();
        }
    }    
}
