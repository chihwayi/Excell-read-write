/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excel;

import java.io.File;

import java.io.FileInputStream;

import java.io.IOException;

import java.sql.*;

import java.util.Locale;

import java.util.logging.Level;

import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFSheet;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import org.apache.poi.ss.usermodel.DataFormatter;

import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.xssf.usermodel.XSSFCell;

import org.apache.poi.xssf.usermodel.XSSFSheet;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author gina
 */
public class excelClass {

    public static void main(String[] args) {

        try {

            Class.forName("com.mysql.jdbc.Driver");

            Connection con = (Connection) DriverManager.getConnection("jdbc:mysql://localhost/exceltest", "root", "");

            con.setAutoCommit(false);

            PreparedStatement pstm = null;

            FileInputStream input = new FileInputStream("/home/gina/NetBeansProjects/WorkinWithExcell/leeTest.xlsx");

            XSSFWorkbook wb = new XSSFWorkbook(input);

            XSSFSheet sheet = wb.getSheetAt(0);

            Row row;

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {

                row = sheet.getRow(i);

                String name = row.getCell(0).getStringCellValue();

                String address = row.getCell(1).getStringCellValue();

                int id = (int) row.getCell(2).getNumericCellValue();

                String sql = "INSERT INTO test VALUES('" + name + "','" + address + "','" + id + "')";

                pstm = (PreparedStatement) con.prepareStatement(sql);

                pstm.execute();

                System.out.println("Import rows " + i);

            }

            con.commit();

            pstm.close();

            con.close();

            input.close();

            System.out.println("Success import excel to mysql table");

        } catch (ClassNotFoundException e) {

            System.out.println(e);

        } catch (SQLException ex) {

            System.out.println(ex);

        } catch (IOException ioe) {

            System.out.println(ioe);

        }

    }
}
