package com.pactera.merge;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

public class MergeService {
    private MergeService(){}
    private static MergeService mergeService=null;
    static MergeService getMergeService(){
        if(mergeService==null){
            mergeService=new MergeService();
        }
        return  mergeService;
    }
    /**
     *  给目标sheet设置title
     * @param targetSheet 目标sheet
     * @param titleFile 获取title的文件
     * @throws IOException IO异常
     */


    void setTargetSheetTitle(XSSFSheet targetSheet, File titleFile) throws IOException {
        InputStream titleIn = new FileInputStream(titleFile);
        XSSFWorkbook titleWB = new XSSFWorkbook(titleIn); //获取源文件的sheet
        XSSFRow titleWBRow = titleWB.getSheetAt(0).getRow(0); //获取第一行

        short firstCellNum = titleWBRow.getFirstCellNum();//获取第一列的列号
        short lastCellNum = titleWBRow.getLastCellNum();//最后一列的列号

        //新sheet没有titile所以不需要获取最后一行
        XSSFRow targetSheetRow = targetSheet.createRow(0);
        while(firstCellNum<lastCellNum){
            XSSFCell targetCell = targetSheetRow.createCell(firstCellNum);
            XSSFCell sourceCell = titleWBRow.getCell(firstCellNum);
            setTargetCellValue(sourceCell,targetCell);
            firstCellNum++;
        }
        System.out.println("title设置成功");
        titleIn.close();
    }
    /**
     *  给目标sheet设置title
     * @param targetSheet 目标sheet
     * @param titleFile 获取title的文件
     * @throws IOException IO异常
     */
    void setTargetSheetTitle(HSSFSheet targetSheet, File titleFile) throws  IOException {
        InputStream titleIn = new FileInputStream(titleFile);
        HSSFWorkbook titleWB = new HSSFWorkbook(titleIn); //获取源文件的sheet
        HSSFRow titleWBRow = titleWB.getSheetAt(0).getRow(0); //获取第一行

        int firstCellNum = (int)titleWBRow.getFirstCellNum();//获取第一列的列号
        int lastCellNum = (int)titleWBRow.getLastCellNum();//最后一列的列号

        //新sheet没有titile所以不需要获取最后一行
        HSSFRow targetSheetRow = targetSheet.createRow(0);
        while(firstCellNum<lastCellNum){
            HSSFCell targetCell = targetSheetRow.createCell(firstCellNum);
            HSSFCell sourceCell = titleWBRow.getCell(firstCellNum);
            setTargetCellValue(sourceCell,targetCell);
            firstCellNum++;
        }
        System.out.println("title设置成功");
        titleIn.close();
    }
   private void setTargetCellValue(HSSFCell sourceCell, HSSFCell targetCell) {
        switch (sourceCell.getCellType()){
            case HSSFCell.CELL_TYPE_STRING: //字符串
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case HSSFCell.CELL_TYPE_BLANK: //空白，空格
                targetCell.setCellValue("");
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN: //布尔值
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case HSSFCell.CELL_TYPE_ERROR: //错误
                targetCell.setCellValue(sourceCell.getErrorCellValue());
                break;
            case HSSFCell.CELL_TYPE_FORMULA: //公式
                //targetCell.setCellValue(sourceCell.getCellFormula()); 获取公式
                targetCell.setCellValue(sourceCell.getCachedFormulaResultType());  //公式的返回值
                break;
            case HSSFCell.CELL_TYPE_NUMERIC: //包含日期处理的情况，可能会出问题
                short format = sourceCell.getCellStyle().getDataFormat();
                SimpleDateFormat sdf ;
                if (format == 14 || format == 31 || format == 57 || format == 58
                        || (176<=format && format<=178) || (182<=format && format<=196)
                        || (210<=format && format<=213) || (208==format ) ) { // 日期
                    sdf = new SimpleDateFormat("MM/dd");
                    double value = sourceCell.getNumericCellValue();
                    Date date = DateUtil.getJavaDate(value);
                    String result = sdf.format(date);
                    targetCell.setCellValue(result);
                } else if (format == 20 || format == 32  || (200<=format && format<=209) ) { // 时间
                    sdf = new SimpleDateFormat("HH:mm");
                    double value = sourceCell.getNumericCellValue();
                    Date date = DateUtil.getJavaDate(value);
                    String result = sdf.format(date);
                    targetCell.setCellValue(result);
                } else { // 不是日期格式
                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                }
                break;
        }
    }
    private void setTargetCellValue(XSSFCell sourceCell, XSSFCell targetCell) {
        switch (sourceCell.getCellType()){
            case XSSFCell.CELL_TYPE_STRING: //字符串
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case XSSFCell.CELL_TYPE_BLANK: //空白，空格
                targetCell.setCellValue("");
                break;
            case XSSFCell.CELL_TYPE_BOOLEAN: //布尔值
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case XSSFCell.CELL_TYPE_ERROR: //错误
                targetCell.setCellValue(sourceCell.getErrorCellString());
                break;
            case XSSFCell.CELL_TYPE_FORMULA: //公式
                //targetCell.setCellValue(sourceCell.getCellFormula()); 获取公式
                targetCell.setCellValue(sourceCell.getCachedFormulaResultType());  //公式的返回值
                break;
            case XSSFCell.CELL_TYPE_NUMERIC: //包含日期处理的情况，可能会出问题
                short format = sourceCell.getCellStyle().getDataFormat();
                SimpleDateFormat sdf ;
                if (format == 14 || format == 31 || format == 57 || format == 58
                        || (176<=format && format<=178) || (182<=format && format<=196)
                        || (210<=format && format<=213) || (208==format ) ) { // 日期
                    sdf = new SimpleDateFormat("MM/dd");
                    double value = sourceCell.getNumericCellValue();
                    Date date = DateUtil.getJavaDate(value);
                    String result = sdf.format(date);
                    targetCell.setCellValue(result);
                } else if (format == 20 || format == 32  || (200<=format && format<=209) ) { // 时间
                    sdf = new SimpleDateFormat("HH:mm");
                    double value = sourceCell.getNumericCellValue();
                    Date date = DateUtil.getJavaDate(value);
                    String result = sdf.format(date);
                    targetCell.setCellValue(result);
                } else { // 不是日期格式
                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                }
                break;
        }
    }
    void copySheet(HSSFSheet sourceSheet, HSSFSheet targetSheet) {
        //总是往目标的最后一行 的 下一行添加数据
        int lastRowNum = targetSheet.getLastRowNum();
        System.out.println("目标sheet的最后一行是："+lastRowNum);
        //数据从第2行复制到最后一行
        int firstNum = sourceSheet.getFirstRowNum()+1;
        int lastNum= sourceSheet.getLastRowNum();
        System.out.println("当前文件的首行是："+(firstNum)+"末行是:"+lastNum+"一共要读取"+(lastNum-firstNum)+"行数据");
        //周报，日报，周数据，bugList中第一行、最后一行的列数相同。
        int lastCellNum = (int)sourceSheet.getRow(firstNum).getLastCellNum();
        for(int i=firstNum;i<=lastNum;i++){
            //内循环结束后，重新初始化单元格的开始位置。
            int firstCellNum = (int)sourceSheet.getRow(i).getFirstCellNum();
            //新建一个行
            System.out.println("正在复制第"+i+"行数据");
            HSSFRow targetRow = targetSheet.createRow(lastRowNum+1);
            for(;firstCellNum<lastCellNum;firstCellNum++){
                HSSFCell sourceCell = sourceSheet.getRow(i).getCell(firstCellNum);
                //在新建的行中，新建列，组成一个单元格
                HSSFCell targetCell = targetRow.createCell(firstCellNum);

                //往目标cell中设置
                if(sourceCell==null){
                    targetCell.setCellValue("0");
                }else{
                    setTargetCellValue(sourceCell,targetCell);
                }

            }
            lastRowNum++;
        }



    }
    void copySheet(XSSFSheet sourceSheet, XSSFSheet targetSheet) {
        //总是往目标的最后一行 的 下一行添加数据
        int lastRowNum = targetSheet.getLastRowNum();
        System.out.println("目标sheet的最后一行是："+lastRowNum);
        //数据从第2行复制到最后一行
        int firstNum = sourceSheet.getFirstRowNum()+1;
        int lastNum= sourceSheet.getLastRowNum();
        System.out.println("当前文件的首行是："+(firstNum)+"末行是:"+lastNum+"一共要读取"+(lastNum-firstNum)+"行数据");
        //周报，日报，周数据，bugList中第一行、最后一行的列数相同。
        short lastCellNum = sourceSheet.getRow(firstNum).getLastCellNum();
        for(int i=firstNum;i<=lastNum;i++){
            //内循环结束后，重新初始化单元格的开始位置。
            short firstCellNum = sourceSheet.getRow(i).getFirstCellNum();
            //新建一个行
            System.out.println("正在复制第"+i+"行数据");
            XSSFRow targetRow = targetSheet.createRow(lastRowNum+1);
            for(;firstCellNum<lastCellNum;firstCellNum++){
                XSSFCell sourceCell = sourceSheet.getRow(i).getCell(firstCellNum);
                //在新建的行中，新建列，组成一个单元格
                XSSFCell targetCell = targetRow.createCell(firstCellNum);
                //往目标cell中设置
                if(sourceCell==null){
                    targetCell.setCellValue("0");
                }else{
                    setTargetCellValue(sourceCell,targetCell);
                }

            }
            lastRowNum++;
        }



    }
}
