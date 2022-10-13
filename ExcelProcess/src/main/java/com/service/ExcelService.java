package com.service;

import com.dao.ExcelDAO;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.annotation.Resource;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedList;

@Service
public class ExcelService {

    @Resource
    private ExcelDAO excelDAO;

    public  void ReadDataFromExcelAndWriteInToDB(String path, String originalFilename){

        Object[][]excel=readExcel(path);
        if (excel==null){
            throw new RuntimeException("No Such Excel");
        }
        excelDAO.ReadDataFromExcelAndWriteInToDB(excel,originalFilename);

    }

    public ArrayList<String> getTableMap(){
       return excelDAO.getTableMap();
    }



    public  Object[][] readExcel(String path) {
        Object[][]objects = null;
        try {
            // 获取文件输入流
            InputStream inputStream = new FileInputStream(path);

            Workbook workbook = null;
            // 截取路径名 . 后面的后缀名，判断是xls还是xlsx

            if (path.substring(path.lastIndexOf(".")+1).equals("xls")){
                workbook = new HSSFWorkbook(inputStream);
            }else if (path.substring(path.lastIndexOf(".")+1).equals("xlsx")){
                workbook = new XSSFWorkbook(inputStream);
            }

            // 获取表
            Sheet sheet = null;
            if (workbook != null) {
                sheet = workbook.getSheetAt(0);
            }else {
                System.out.println("sheet==null");
                return null;
            }
            // 获取总的列数
            int rowNum=0;
            for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                // 循环读取每一个格
                Row row = sheet.getRow(i);
                rowNum=Math.max(row.getPhysicalNumberOfCells(),rowNum);
            }

            // and循环读取每一行

            objects=new Object[sheet.getPhysicalNumberOfRows()]
                    [rowNum];


            for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                // 循环读取每一个格
                Row row = sheet.getRow(i);
                // row.getPhysicalNumberOfCells()获取总的列数
                for (int index = 0; index < row.getPhysicalNumberOfCells(); index++) {
                    // 获取数据

                    Cell cell = row.getCell(index);

                    // 转换为字符串类型
                    if (cell!=null) {
                        cell.setCellType(CellType.STRING);
                        // 获取得到字符串
                        String id=cell.getStringCellValue();

                        objects[i][index]=id;
                    }else {
                        objects[i][index]=null;
                    }
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        return objects;
    }


    public HSSFWorkbook createExcel(String excelName) {

        HSSFWorkbook workbook = new HSSFWorkbook();

        excelDAO.writeDataIntoExcel(workbook,excelName);


        return workbook;


    }





}
