/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.ingos.exeltemplatelib;

import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.ss.util.CellReference;
import teamworks.TWList;
import teamworks.TWObject;

/**
 *
 * @author aarapov
 */
public class MSFactory {
    
    public ByteArrayOutputStream formXLSDocument(ArrayList<HashMap<String, String>> serviceTaskList) throws FileNotFoundException, IOException {
        HSSFWorkbook myExcelBook;
        
        myExcelBook = new HSSFWorkbook(this.getClass().getClassLoader().getResourceAsStream("template.xls"));
        
        HSSFSheet myExcelSheet = myExcelBook.getSheetAt(0);
        
        HSSFCellStyle style = myExcelBook.createCellStyle();
        Font font = myExcelBook.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Times New Roman");
        style.setFont(font);
        
        int column;
        int row = 18;
        int lastRow = myExcelSheet.getLastRowNum();
        String value;
        
        for (int i = 0; i < serviceTaskList.size(); i++) {
            HashMap<String, String> keyMap = serviceTaskList.get(i);
            //userName ФИО
            column = CellReference.convertColStringToIndex("A");
            value = keyMap.get("userName");
            myExcelSheet.createRow(row).createCell(column).setCellValue(value);

            //department Структурное подразделение
            column = CellReference.convertColStringToIndex("C");
            value = keyMap.get("department");
            myExcelSheet.getRow(row).createCell(column).setCellValue(value);

            //position Должность
            column = CellReference.convertColStringToIndex("D");
            value = keyMap.get("position");
            myExcelSheet.getRow(row).createCell(column).setCellValue(value);

            //country+town Страна, город
            column = CellReference.convertColStringToIndex("E");
            value = keyMap.get("country") + ", " + keyMap.get("town");
            myExcelSheet.getRow(row).createCell(column).setCellValue(value);

            //destination Место назначения
            column = CellReference.convertColStringToIndex("F");
            value = keyMap.get("destination");
            myExcelSheet.getRow(row).createCell(column).setCellValue(value);

            //startDate Дата начала
            column = CellReference.convertColStringToIndex("G");
            value = keyMap.get("startDate");
            myExcelSheet.getRow(row).createCell(column).setCellValue(value);

            //endDate Дата конца
            column = CellReference.convertColStringToIndex("H");
            value = keyMap.get("endDate");
            myExcelSheet.getRow(row).createCell(column, Cell.CELL_TYPE_STRING).setCellValue(value);

            //term Срок
            column = CellReference.convertColStringToIndex("I");
            value = keyMap.get("term");
            myExcelSheet.getRow(row).createCell(column).setCellValue(value);

            //target Цель
            column = CellReference.convertColStringToIndex("J");
            value = keyMap.get("target");
            myExcelSheet.getRow(row).createCell(column).setCellValue(value);

            //добавить границы новой строки A - L
            addAllBordersToRow(myExcelSheet.getRow(row), 0, 12, style);
            
            row++;
            lastRow++;
            myExcelSheet.shiftRows(row, lastRow, 1, true, true, true);
        }
        
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        myExcelBook.write(os);
        
        return os;
    }
    
    public String getbase64XLSBytes(TWList list) throws IOException {
        
        ArrayList<HashMap<String, String>> arrayList = new ArrayList<HashMap<String, String>>();
        
        for (int i = 0; i < list.getArraySize(); i++) {
            Object obj = list.getArrayData(i);
            HashMap<String, String> map = new HashMap<String, String>();
            if (obj instanceof TWObject) {
                for (String name : ((TWObject) obj).getPropertyNames()) {
                    Object value = ((TWObject) obj).getPropertyValue(name);
                    if (value instanceof Date) {
                        DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
                        map.put(name, df.format((Date) value));
                    } else if (value instanceof GregorianCalendar) {
                        SimpleDateFormat sd = new SimpleDateFormat("dd/MM/yyyy");
                        GregorianCalendar cald = (GregorianCalendar) value;
                        sd.setCalendar(cald);
                        map.put(name, sd.format(cald.getTime()));
                    } else {
                        map.put(name, value.toString());
                    }
                }
            }
            arrayList.add(map);
        }
        String base64 = new sun.misc.BASE64Encoder().encode(formXLSDocument(arrayList).toByteArray());;
        return base64;
        
    }
    
    private void addAllBordersToRow(Row row, int start, int end, HSSFCellStyle style) {
        
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        
        for (int i = start; i < end; i++) {
            Cell cell = row.getCell(i);
            if (cell == null) {
                cell = row.createCell(i);
            }
            cell.setCellStyle(style);
        }
    }

//    private void writeToDatabase(InputStream inStream) throws NamingException, SQLException {
//        Context ctx = new InitialContext();
//        DataSource ds = (DataSource) ctx.lookup("jdbc/AIS_DS");
//        Connection con = ds.getConnection();
//
//        String sql = "insert into LSW_BPD_INSTANCE_DOCUMENTS ("
//                + "DOC_ID,"
//                + "BPD_INSTANCE_ID,"
//                + "PARENT_DOC_ID, "
//                + "DOC_NAME, "
//                + "FILE_NAME, "
//                + "FILE_TYPE, "
//                + "AUTHOR_ID, "
//                + "AUTHORED_DTG,"
//                + "VERSION,"
//                + "MIGRATION_STATE,"
//                + "DOCUMENT) "
//                + "values (?,?,?,?,?,?,?,?,?,?,?)";
//        PreparedStatement pre = con.prepareStatement(sql);
//        pre.setInt(1, 0);
//        pre.setInt(2, 0);
//        pre.setInt(3, 0);
//        pre.setNString(4, "0");
//        pre.setNString(5, "0");
//        pre.setInt(6, 0);
//        pre.setInt(7, 0);
//        pre.setDate(8, new Date(System.currentTimeMillis()));
//        pre.setInt(9, 0);
//        pre.setInt(10, 0);
//        pre.setBinaryStream(11, inStream);
//
//        pre.executeUpdate();
//    }
}
