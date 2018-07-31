package com.pactera.merge;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;



public class MergeAction  {
private static MergeService mergeService =MergeService.getMergeService();
    public XSSFWorkbook mergeXSSF(XSSFWorkbook targetWB, String targetSheetName,String ...path)  {
        XSSFSheet targetSheet = targetWB.createSheet(targetSheetName);
        //需要获取第一个文件的title
        File titleFile = new File(path[0]);
        System.out.println("开始设置title"+titleFile.getName());
        try {
           mergeService.setTargetSheetTitle(targetSheet,titleFile);
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

    public HSSFWorkbook mergeHSSF(HSSFWorkbook targetWB, String targetSheetName, String ... path)  {
        HSSFSheet targetSheet = targetWB.createSheet(targetSheetName);
        //需要获取第一个文件的title
        File titleFile = new File(path[0]);
        System.out.println("开始设置title"+titleFile.getName());
        try {
            mergeService.setTargetSheetTitle(targetSheet,titleFile);
            System.out.println("开始读取所有文件");
            for(String p :path){
                //读入xls文件
                File file = new File(p);
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
}



