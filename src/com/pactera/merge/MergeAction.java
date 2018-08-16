package com.pactera.merge;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;

public class MergeAction {
    private MergePlus mergePlus =new MergePlus();
    /**
     * 将该excel下的所有sheet并合并到一个文件的同一个sheet中
     * @param targetWorkbook 目标excel文件
     * @param sourceFiles 要复制的所有文件
     * @return 目标excel文件
     * @throws IOException 文件处理时的IO异常
     * @throws FileReadException  文件获取时的异常
     */
    public Workbook copyAllToOneSheet(Workbook targetWorkbook, File[] sourceFiles) throws IOException, FileReadException {
        Sheet targetSheet = targetWorkbook.createSheet("11");
        File sourceFile = sourceFiles[0];
        //给目标sheet设置title
        mergePlus.copyFirstRow(targetSheet,sourceFile,0);
        for(File file :sourceFiles){
            System.out.println("当前读取的文件是："+file.getName());
            Workbook sourceWorkbook = mergePlus.excelEndWith(file,1);
            int num = sourceWorkbook.getNumberOfSheets();
            while(num>0){
                Sheet sourceSheet = sourceWorkbook.getSheetAt(num - 1);
                if(sourceSheet.getLastRowNum()>0){
                    mergePlus.copySheet(sourceSheet,targetSheet);
                }
                num--;
            }
        }
        return targetWorkbook;
    }
    public Workbook mergeLastSheet(Workbook targetWorkbook, File[] sourceFiles)throws IOException,FileReadException{
        Sheet targetSheet = targetWorkbook.createSheet();
        Workbook workbook = mergePlus.excelEndWith(sourceFiles[0], 1);
        int num = workbook.getNumberOfSheets();
        Sheet sheet = workbook.getSheetAt(num - 1);
        while(sheet.getLastRowNum()<=0){
            num--;
            sheet=workbook.getSheetAt(num-1);
        }
        if(num>0){
            mergePlus.copyFirstRow(sheet,targetSheet);
        }

        for(File files:sourceFiles){
            Workbook sourceWB = mergePlus.excelEndWith(files, 1);
            int number = sourceWB.getNumberOfSheets();
            Sheet sourceSheet = sourceWB.getSheetAt(number - 1);
            while(sourceSheet.getLastRowNum()<=0){
                number--;
                //最后一个sheet为空的话，更新sourceSheet
                sourceSheet=sourceWB.getSheetAt(number-1);
            }
            if(number>0){
                mergePlus.copySheet(sourceSheet, targetSheet);
            }else{
                throw new FileReadException("要读取的excel文件不能为空！！！");
            }
        }
        return targetWorkbook;


    }
}
