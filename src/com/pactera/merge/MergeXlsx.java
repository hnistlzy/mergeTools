package com.pactera.merge;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;


public class MergeXlsx  {

    public XSSFWorkbook readFiles(XSSFWorkbook targetWB, String targetSheetName,String ...path)  {
        XSSFSheet targetSheet = targetWB.createSheet(targetSheetName);
        //需要获取第一个文件的title
        File titleFile = new File(path[0]);
        System.out.println("开始设置title"+titleFile.getName());
        try {
            setTargetSheetTitle(targetSheet,titleFile);
            System.out.println("开始读取所有文件");
        for(String p :path){
            //读入xls文件
            File file = new File(p);
            System.out.println("当前读取的文件是:"+file.getName());
            InputStream in = new FileInputStream(file);
            XSSFWorkbook sourceFile = new XSSFWorkbook(in);
            //当前excel文件中sheet的数量，每个sheet都要导入
            int num = sourceFile.getNumberOfSheets();
            while (num>0){
                copySheet(sourceFile.getSheetAt(num - 1), targetSheet);
                num--;
            }
            in.close();
        }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return targetWB;
    }
    /**
     *  给目标sheet设置title
     * @param targetSheet 目标sheet
     * @param titleFile 获取title的文件
     * @throws IOException IO异常
     */
    private void setTargetSheetTitle(XSSFSheet targetSheet, File titleFile) throws  IOException {
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
     * 判断源单元格的类型，并往新的单元格中存放数据
     * @param sourceCell 源单元格
     * @param targetCell 目标单元格
     */
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

    /**
     *  将sheet的内容追加到现有sheet中
     * @param sourceSheet 源sheet
     * @param targetSheet 目标sheet
     */
    private void copySheet(XSSFSheet sourceSheet, XSSFSheet targetSheet) {
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



    }}


