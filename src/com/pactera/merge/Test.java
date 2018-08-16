package com.pactera.merge;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;

public class Test {
    /**
     * 为什么设置的样式没有生效？
     */
    public static void main(String[] args) {
        XSSFWorkbook excel = new XSSFWorkbook();
        XSSFSheet sheet = excel.createSheet("测试");

        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell(0);
        cell.setCellValue("测试样式");

        XSSFCellStyle cellStyle = excel.createCellStyle();
        cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_DASH_DOT);
        cellStyle.setBorderTop(XSSFCellStyle.BORDER_DASH_DOT);
        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cellStyle.setBottomBorderColor((short)1);

        File file = new File("D:\\2.xlsx");
        try{
            FileOutputStream fileOutputStream = new FileOutputStream(file);
            excel.write(fileOutputStream);
            fileOutputStream.close();
        }catch (Exception e){
            e.printStackTrace();
        }


    }
}
