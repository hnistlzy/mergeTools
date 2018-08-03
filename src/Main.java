import com.pactera.merge.MergeAction;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;


public class Main {
    public static void main(String[] args) {
        MergeAction mergeAction = new MergeAction();
        //文件源路径
        File fileD = new File("D:\\周报");
        File[] fileList = fileD.listFiles();
        FileOutputStream fileOutputStream;

        File resultFile=null;
        if(fileList!=null&&fileList.length>0){
            resultFile= mergeAction.judgeFile(fileList, "D:\\周报\\bugList");//输出路径
        }
        try{
            if(resultFile!=null){
                fileOutputStream = new FileOutputStream(resultFile);
                if(!resultFile.getName().contains(".xlsx")){
                    HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
                    HSSFWorkbook bug = mergeAction.mergeHSSF(hssfWorkbook, "bug", fileList);
                    bug.write(fileOutputStream);
                    fileOutputStream.close();
                }else {
                    XSSFWorkbook xssfSheets = new XSSFWorkbook();
                    XSSFWorkbook bug = mergeAction.mergeXSSF(xssfSheets, "bug", fileList);
                    bug.write(fileOutputStream);
                    fileOutputStream.close();
                }
            }
        }catch (IOException e){
            e.printStackTrace();
        }

    }
}
