// Copyright [c] 2002 Artemis International Solutions Corporation
package artemispm.jvdo2;
// Updating this file online...
// from machine - second changes

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Types;

import oracle.jdbc.OracleTypes;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import artemispm.jvdo.JVJResultSet;
import artemispm.serverutil.AppMgr;
import artemispm.utility.ApmRuntimeException;
import artemispm.utility.VersionHolder;

/**
 * sQL Factory for sPools.
 * <P>
 *
 * @author softserve
 */
public class HelloWorld extends JVBaseSQL {
    // csize of the sql column in av_xlreportsql
    private static final int VARCHAR4000 = 4000;

    /**
     * Constructor
     */
    public JVStoredSelSQL(String user, String dataset, String locale) throws SQLException {
        // Method implementation
        super(user, dataset, locale);
    }

    public JVSet fetch(Connection con, String[] params) throws SQLException {
        // Method implementation
        if (params.length > 0) {
            if ("delete".equals(params[0])) {
                doDelete(con, params[1]);
                return new JVStoredSelSet();
            } else if ("test".equals(params[0])) {
                JVStoredSel sel;
                try {
                    sel = doTest(con, params[1]);
                    if (sel != null) {
                        JVStoredSelSet selSet = new JVStoredSelSet();
                        selSet.add(sel);
                        return selSet;
                    }
                } catch (IOException e) {
                    // Pending Auto-generated catch block
                    VersionHolder.prntStkTrc(VersionHolder.CONTEXT, e);
                }
                return new JVStoredSelSet();
            }
        }
        return getStored(con);
    }

    private void doDelete(Connection con, String name) throws SQLException {
        String sql = "delete from av_xlreportsql where sqlid='" + name + "' and orderno=1";
        PreparedStatement stmt = con.prepareStatement(sql);
        // stmt.setString[1, name]
        stmt.executeUpdate();
        con.commit();
        stmt.close();
        stmt = null;
    }

    private JVStoredSel doTest(Connection con, String sql) throws IOException {
        PreparedStatement stmt;
        ResultSet oRecords;
        try {
            //stmt = con.prepareStatement(sql);
        	
        	
        	CallableStatement proc = null;                
        	   proc = con.prepareCall("{ call viewsreportsp(?, ?, ?) }");
        	   proc.setString(1, "Name");
        	   proc.setInt(2, 1);
        	   proc.registerOutParameter(3, OracleTypes.CURSOR);
        	   proc.execute();
        	   
        	   oRecords = (ResultSet)proc.getObject(3);

        	//PreparedStatement ps=con.prepareStatement(sql);
            //oRecords = ps.executeQuery();
            ResultSetMetaData meta = oRecords.getMetaData();
            int cols = meta.getColumnCount();
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.createSheet("sheet1");
            HSSFRow row = sheet.createRow((short) 0);
            boolean[] numeric = new boolean[cols + 1];
            boolean[] floating = new boolean[cols + 1];
            boolean[] date = new boolean[cols + 1];
            for (int i = 1; i <= cols; i++) {
                row.createCell((short) (i - 1)).setCellValue(meta.getColumnName(i));
                int type = meta.getColumnType(i);
                numeric[i] = (type == Types.DECIMAL && meta.getScale(i) == 0);
                floating[i] = !numeric[i] && (type == Types.DECIMAL || type == Types.REAL || type == Types.DOUBLE || type == Types.FLOAT);
                date[i] = (type == Types.DATE || type == Types.TIMESTAMP);
            }
            
            /*
            PreparedStatement ps=con.prepareStatement(result);
            ResultSet rs=ps.executeQuery();
           
            ResultSetMetaData meta=rs.getMetaData();
                    
            int cols=meta.getColumnCount();
            
            HSSFWorkbook wb          = new HSSFWorkbook();
            HSSFSheet sheet = wb.createSheet("sheet1");
         
          
            HSSFRow row     = sheet.createRow((short)0); 
               
            boolean numeric[]=new boolean[cols+1];
            boolean floating[]=new boolean[cols+1];
            boolean date[]=new boolean[cols+1];
                            
            for (int i=1;i<=cols;i++) {
                row.createCell((short)(i-1)).setCellValue(meta.getColumnName(i));
                               
                int type=meta.getColumnType(i);
                numeric[i]=(type==Types.DECIMAL
                            &&  meta.getScale(i)==0);
                floating[i]=!numeric[i] && 
                     (type==Types.DECIMAL || type==Types.REAL
                        || type==Types.DOUBLE || type==Types.FLOAT);
                date[i]=(type==Types.DATE || type==Types.TIMESTAMP);
            }
                    
           
         
            short irows=0;
                   
            while(rs.next()) {
                irows++;
                row     = sheet.createRow(irows);
                for (int i=1;i<=cols;i++) {
                    HSSFCell cell = row.createCell((short)(i-1));
                   // cell.setEncoding( HSSFCell.ENCODING_COMPRESSED_UNICODE );
                    if (date[i]) cell.setCellValue("");
                           else if (numeric[i]) cell.setCellValue(rs.getLong(i));
                           else if (floating[i]) cell.setCellValue(rs.getDouble(i));
                           else cell.setCellValue(rs.getString(i));
                      } 
            }
          
          String wfn="A7test"+System.currentTimeMillis()+".xls";
          File workbook=new File("C:\\TEMP"+File.separator + wfn);
          
          FileOutputStream fos=new FileOutputStream(workbook);
          wb.write(fos);
          fos.close();
//          downloadFile("",wfn, workbook ,"*");
          workbook.delete();  
          */
            
            
            short irows = 0;
            while (oRecords.next()) {
                irows++;
                row = sheet.createRow(irows);
                for (int i = 1; i <= cols; i++) {
                    HSSFCell cell = row.createCell((short) (i - 1));
                    // cell.setEncoding[ HSSFCell.ENCODING_COMPRESSED_UNICODE ]
                    if (date[i]) {
                        cell.setCellValue(JVBaseSQL.getPossNullDate((JVJResultSet) oRecords, i).toISOString());
                    } else if (numeric[i]) {
                        cell.setCellValue(oRecords.getLong(i));
                    } else if (floating[i]) {
                        cell.setCellValue(oRecords.getDouble(i));
                    } else {
                        cell.setCellValue(oRecords.getString(i));
                    }
                }
            }
            String wfn = "VIEWStest" + System.currentTimeMillis() + ".xls";
            String rwfn = wfn;
            File workbook = new File(AppMgr.getInstance().getWorkDir() + File.separator + wfn);
            FileOutputStream fos = new FileOutputStream(workbook);
            wb.write(fos);
            fos.close();
            workbook.delete();
            // open xls
            FileOutputStream foutput;
            String tmpDir = artemispm.serverutil.AppMgr.getTempDir();
            if (!tmpDir.endsWith(File.separator)) {
                tmpDir += File.separator;
            }
            wfn = tmpDir + wfn;
            foutput = new FileOutputStream(wfn);
            wb.write(foutput);
            foutput.close();
            //stmt.close();
            //ps.close();
            stmt = null;
            JVStoredSel sel = new JVStoredSel();
            sel.setSql(rwfn);
            return sel;
        } catch (SQLException e) {
            JVStoredSel sel = new JVStoredSel();
            sel.setSql(e.getMessage());
            VersionHolder.prntStkTrc(VersionHolder.CONTEXT, e);
            return sel;
        }
    }

    public JVObject fetchObject(Connection con, String[] params) throws SQLException {
        // Method implementation
        // we expect the name
        return getStored(con, params[0]);
    }

    public JVStoredSelSet getStored(Connection con) throws SQLException {
        // Method implementation
        JVStoredSelSet set = new JVStoredSelSet();
        getStoredSQL(con, null, set);
        return set;
    }

    public JVStoredSel getStored(Connection con, String name) throws SQLException {
        // Method implementation
        return getStoredSQL(con, name, null);
    }

    private JVStoredSel getStoredSQL(Connection conn, String name, JVSet set) throws SQLException {
        // this is where we do the work
        JVStoredSel storedsel = null;
        ResultSet oRecords = null;
        String sql;
        int col;
        if (set != null && !set.isEmpty()) {
            throw new ApmRuntimeException("Collections cannot have data when passed to fetch methods");
        }
        sql = "select sqlid, sql from av_xlreportsql";
        // logDebug[sq]
        try {
            PreparedStatement stmt = conn.prepareStatement(sql);
            if (name != null) {
                stmt.setString(1, rewrapQuotes(name));
            }
            oRecords = stmt.executeQuery();
            while (oRecords.next()) {
                col = 1;
                storedsel = new JVStoredSel();
                storedsel.setStoredName(oRecords.getString(col++));
                storedsel.setSql(oRecords.getString(col++));
                if (set != null) {
                    set.addItem(storedsel);
                }
                storedsel.setSaved();
            }
            if (storedsel != null) {
                storedsel.setSaved();
            }
            stmt.close();
            stmt = null;
        } finally {
            try {
                oRecords.close();
                oRecords = null;
            } catch (SQLException e) {
                VersionHolder.prntStkTrc(VersionHolder.CONTEXT, e);
            }
        }
        return storedsel;
    }

    private void insertStored(Connection con, JVStoredSel stored) throws SQLException {
        StringBuilder sql = new StringBuilder(SQL_INITIALLENGTH);
        PreparedStatement stmt;
        if (!isValidForSave(stored)) {
            return;
        }
        sql.append("insert into av_xlreportsql (sqlid,orderno,sql) values(?,?,?)");
        stmt = con.prepareStatement(sql.toString());
        setReportFields(stored, stmt);
        stmt.executeUpdate();
        stmt.close();
    }

    private int setReportFields(JVStoredSel stored, PreparedStatement stmt) {
        int fieldNo = 0;
        int maxlen = VARCHAR4000;
        String sql = stored.getSql();
        if (!isValidForSave(stored)) {
            return 0;
        }
        try {
            int emitted = 0;
            for (int orderno = 1; emitted < sql.length(); orderno++) {
                setString(stmt, 1, stored.getStoredName());
                setPossNullInt(stmt, 2, orderno);
                int emit = sql.length() - emitted;
                if (emit > maxlen) {
                    emit = maxlen;
                }
                setString(stmt, 3, sql.substring(emitted, emitted + emit));
                emitted += emit;
            }
        } catch (SQLException e) {
            VersionHolder.prntStkTrc(VersionHolder.CONTEXT, e);
        }
        return fieldNo;
    }

    public void save(Connection con, JVSet set) throws SQLException, JVException {
        // Method implementation
    }

    public void saveObject(Connection con, JVObject obj) throws SQLException, JVException {
        // Method implementation
        saveStored(con, (JVStoredSel) obj);
    }

    private void saveStored(Connection con, JVStoredSel stored) throws SQLException {
        insertStored(con, stored);
    }
}
