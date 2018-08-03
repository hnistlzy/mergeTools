package com.pactera.contrast;

import com.pactera.pojo.User;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class SelectWayImp implements SelectBehavior {
    private User user ;
    @Override
    public void selectWayXlsx(File file, String str) throws IOException {
        if(file.getName().contains("周数据")&&str.equals("周数据")){
            System.out.println("统计周数据");
            FileInputStream  in = new FileInputStream(file);
            XSSFWorkbook  wb= new XSSFWorkbook(in);
            int num = wb.getNumberOfSheets();
            getWeeklyValue(wb,1,6);
            //通过反射获得属性名
            //map<string,string>  <属性名，值>
            //list<map>
            //user.setvalue(map.getvalue())
            while(num>0){
                XSSFSheet sheet = wb.getSheetAt(num);
                Iterator<Row> itr = sheet.iterator();
//                while(itr.hasNext()){
//                    User user = new User();
//                    Row next = itr.next();
//
//                }
//                int firstRowNum = sheetAt.getFirstRowNum();
//                int lastRowNum = sheetAt.getLastRowNum();
//                for(;firstRowNum<=lastRowNum;firstRowNum++){
//                    XSSFRow row = sheetAt.getRow(firstRowNum);
//                    int firstCellNum = (int)row.getFirstCellNum();
//                    int lastCellNum = (int)row.getLastCellNum();
//                    for(;firstCellNum<=lastCellNum;firstCellNum++){
//                        user =new User();
//                        XSSFCell cell = row.getCell(firstCellNum);
//                        getCellValue(cell);
//                    }
//                }
                num--;
            }
        }else if(file.getName().contains("BugList")&& str.equals("BugList")){
            System.out.println("统计BugList中的数据");
        }
    }

    private void getWeeklyValue(XSSFWorkbook wb, int first, int last) {
    }

    private void getCellValue(XSSFCell cell) {
        switch (cell.getCellType()){
            case XSSFCell.CELL_TYPE_BLANK:
                break;
            case XSSFCell.CELL_TYPE_BOOLEAN:
                cell.getBooleanCellValue();
                break;
        }
    }

    @Override
    public void selectWayXls(File file, String str) {

    }
}
