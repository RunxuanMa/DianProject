package com.dao;



import com.Bean.Excel;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.dao.DataIntegrityViolationException;
import org.springframework.data.relational.core.sql.In;
import org.springframework.jdbc.core.BeanPropertyRowMapper;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Repository;

import javax.annotation.Resource;
import java.util.*;

@Repository(value = "ExcelDAO")
public class ExcelDAO {

    @Resource
    private Excel excel;

    @Resource
    private  JdbcTemplate jdbcTemplate;

    public ArrayList<String> getTableMap(){

        return  (ArrayList<String>) jdbcTemplate.queryForList
                (" select `filename` from `map` ", String.class);
    }



    public void ReadDataFromExcelAndWriteInToDB(Object[][] excelData, String originalFilename) {


        String tableName="excel"+System.currentTimeMillis();

//        String[] split = originalFilename.split("\\.");
//
//        originalFilename=split[0];
        List<String> tablename = jdbcTemplate.queryForList(" select `tablename` from map where `filename`=" + "'" + originalFilename + "'",
                String.class
        );

        if (tablename.size()>0){
            throw new RuntimeException("警告，文件重名！");
        }
        jdbcTemplate.update(" INSERT INTO `map`  values(?,?) ",tableName,originalFilename);



        String createTableSql="CREATE TABLE `"+tableName+"`  (\n" +
                "  `col_1` VARCHAR(32) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL,\n" +
                "  `col_2` VARCHAR(32) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL,\n" +
                "  `col_3` DOUBLE NULL DEFAULT NULL,\n" +
                "  `col_4` VARCHAR(32) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL,\n" +
                "  `col_5` DOUBLE NULL DEFAULT NULL\n" +
                ") ENGINE = INNODB CHARACTER SET = utf8 COLLATE=utf8_general_ci ROW_FORMAT = DYNAMIC;";





        jdbcTemplate.update(createTableSql);


        int[][]toLargeData=new int[6][excelData.length];
        int[][]nullData=new int[6][excelData.length];
        StringBuilder exceptionStr = new StringBuilder();
        for (int i=0;i<excelData.length;i++) {
            Object[] excelDatum = excelData[i];

            if (excelDatum[0] != null && ((String) excelDatum[0]).length() > 32) {
                toLargeData[1][i] = 1;
            }
            if (excelDatum[1] != null && ((String) excelDatum[1]).length() > 32) {
                toLargeData[2][i] = 1;
            }
            if (excelDatum[4] != null && ((String) excelDatum[4]).length() > 32) {
                toLargeData[5][i] = 1;
            }

            if (excelDatum[0]==null||excelDatum[0] == "") {
                nullData[1][i]=1;
            }
            if (excelDatum[1]==null||excelDatum[1] == "") {
                nullData[2][i]=1;
            }
            if (excelDatum[2]==null||excelDatum[2] == "") {
                nullData[3][i]=1;
            }



        }
        for (int k = 1; k <= 3; k++) {

                for (int j = 0; j < excelData.length; j++) {

                    if (nullData[k][j]==1) {
                        exceptionStr.append("文件的第").append(k).append("列").append("第").append(1+j).append("行为空,这是必填项!\n");
                    }
                }

        }

        for (int k = 1; k < 6; k++) {
            for (int j = 0; j < excelData.length; j++) {
                if (toLargeData[k][j] == 1) {
                    exceptionStr.append("文件的第").append(k).append("列").append("第").append(j+1).append("行数据过大!\n");
                }

            }
        }
        for (int i=0;i<excelData.length;i++) {
            Object[] excelDatum = excelData[i];
            try {
                jdbcTemplate.update("insert into " + tableName + " values(?,?,?,?,?)",
                        excelDatum[0],
                        excelDatum[1],
                        excelDatum[2],
                        excelDatum[3],
                        excelDatum[4]);
            } catch (DataIntegrityViolationException e) {

                String msg = e.getMessage();

                msg = msg.split("'")[1];

                if (msg.contains("3") && i == 0) { //特殊情况
                } else {
                    jdbcTemplate.update(" delete from map where filename='" + originalFilename + "'");
                    jdbcTemplate.update("drop table " + tableName);
                    //  exceptionStr+="您的文件存在过长字段！过长字段位于" + msg + ",在第" + (i + 1) + "行";
                    throw new RuntimeException(exceptionStr.toString());
                }
            }
        }
            if (!exceptionStr.toString().equals("")){
                jdbcTemplate.update(" delete from map where filename='"+originalFilename+"'");
                jdbcTemplate.update("drop table "+tableName);
                throw new RuntimeException(exceptionStr.toString());
            }

        }



    public void writeDataIntoExcel(HSSFWorkbook workbook, String excelName){
        String tableName="";

        try {
            tableName = jdbcTemplate.queryForObject(" select `tablename` from map where `filename`=" +"'"+excelName+"'"
                    , String.class);
        }catch (Exception e){
            throw new RuntimeException("无法找到您想要的数据！");
        }


        BeanPropertyRowMapper<Excel> rowMapper = new BeanPropertyRowMapper<>(Excel.class);


        List<Excel> query = jdbcTemplate.query(" SELECT * FROM "+tableName,rowMapper);

//        for (Excel excel:query){
//            System.out.println(excel.toString());
//        }


        HSSFSheet sheet = workbook.createSheet("0");
        int row=0;

        for (Excel excels:query){
            int num=0;
            HSSFRow row1 = sheet.createRow(row);
            row++;
            String col_1 = excels.getCol_1();
            String col_2 = excels.getCol_2();
            Double col_3 = excels.getCol_3();
            String col_4 = excels.getCol_4();
            Double col_5 = excels.getCol_5();

            if (col_1!=null) {
                row1.createCell(num++).setCellValue(col_1);
            }
            if (col_2!=null) {
                row1.createCell(num++).setCellValue(col_2);
            }
            if (col_3!=null) {
                row1.createCell(num++).setCellValue(col_3);
            }
            if (col_4!=null) {
                row1.createCell(num++).setCellValue(col_4);
            }
            if (col_5!=null) {
                row1.createCell(num++).setCellValue(col_5);
            }

        }


    }




}
