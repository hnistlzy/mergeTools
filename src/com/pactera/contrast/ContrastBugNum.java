package com.pactera.contrast;

import java.io.File;
import java.io.IOException;


public class ContrastBugNum {
    //读取文件
    //判断文件类型 1.xlsx 2.xls
    //判断文件名 是否包含字段 k
         //如果这个字段是buglist
              //1.统计该文件中第i列的出现次数。 返回姓名 跟 出现次数
        //如果这个字段是周数据
              //1.获取该文件中第i列的值，并相加。 返回姓名 跟 出现次数。
    //比较两个返回结果
    private SelectBehavior selectBehavior;
    public void readFiles(String path,String str) throws IOException {
        File file = new File(path);
        String fileName = file.getName();
        String fileType;
        if(fileName.contains(".xlsx")){
            System.out.println("使用XSSF来读取数据");
            selectBehavior.selectWayXlsx(file,str);
        }else {
            System.out.println("使用HSSF来读取数据");
            selectBehavior.selectWayXls(file,str);
        }
    }

    private void selectWay(File file, String str) {
        if(file.getName().contains(str) && str.equals("周数据") ){
            System.out.println("个人周数据");
        }
    }


}


