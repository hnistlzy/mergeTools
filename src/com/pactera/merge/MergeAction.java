package com.pactera.merge;

import org.apache.poi.hslf.model.Sheet;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;



public class MergeAction  {
private static MergeService mergeService =MergeService.getMergeService();

    public XSSFWorkbook mergeXSSF(XSSFWorkbook targetWB, String targetSheetName,File[] files)  {
        XSSFSheet targetSheet = targetWB.createSheet(targetSheetName);
        //需要获取第一个文件的title
        File titleFile = files[0];
        System.out.println("开始设置title"+titleFile.getName());
        try {
           mergeService.setTargetSheetTitle(targetSheet,titleFile);
            System.out.println("开始读取所有文件");
        for(File file :files){
            //读入xls文件
            System.out.println("当前读取的文件是:"+file.getName());
            InputStream in = new FileInputStream(file);
            XSSFWorkbook sourceFile = new XSSFWorkbook(in);
            //当前excel文件中sheet的数量，每个sheet都要导入

            for(XSSFSheet sheet:sourceFile){
                mergeService.copySheet(sheet,targetSheet,sourceFile.getCreationHelper().createFormulaEvaluator());
            }
            in.close();
        }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return targetWB;
    }

    public HSSFWorkbook mergeHSSF(HSSFWorkbook targetWB, String targetSheetName, File[] files)  {
        HSSFSheet targetSheet = targetWB.createSheet(targetSheetName);
        //需要获取第一个文件的title
         File titleFile = files[0];
        System.out.println("开始设置title"+titleFile.getName());
        try {
            mergeService.setTargetSheetTitle(targetSheet,titleFile);
            System.out.println("开始读取所有文件");
            for(File file :files){
                //读入xls文件
                System.out.println("当前读取的文件是:"+file.getName());
                InputStream in = new FileInputStream(file);
                HSSFWorkbook sourceFile = new HSSFWorkbook(in);
                //当前excel文件中sheet的数量，每个sheet都要导入
                int num = sourceFile.getNumberOfSheets();
                while (num>0){
                    mergeService.copySheet(sourceFile.getSheetAt(num - 1), targetSheet);
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
     *  根据获取到的第一个文件的后缀名判断要创建哪种类型的文件
     * @param files file数组
     * @param targetPath  合并结果输出目录 比如;D:\\周报
     * @return 用来储存合并结果的文件
     */
    public File judgeFile(File[] files,String targetPath){
        File file;
        String str=files[0].getName();
        StringBuilder sb = new StringBuilder(targetPath);
        if(str.contains(".xlsx")){
            file = new File(sb.append(".xlsx").toString());
        }else{
            file = new File(sb.append(".xls").toString());
        }
        return  file;
    }
}



