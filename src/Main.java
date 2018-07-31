import com.pactera.merge.MergeAction;
import com.pactera.merge.MergeService;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;


public class Main {
    /**
     * 默认认为输入的这一组数据都为同一个格式
     * @param path  地址数组
     */
    private int readFiles(String ...path){
        int status=0;
        String str=path[0];
        if(str.contains(".xlsx")){
            status=1;
        }
        return  status;
}
    public static void main(String[] args) {
//        MergeXlsx mergeXls = new MergeXlsx();
//        XSSFWorkbook targetWb = new XSSFWorkbook();
//        XSSFWorkbook resultWB = mergeXls.mergeXSSF(targetWb, "buglist", "D:\\升级.xlsx", "D:\\麻将.xlsx");
//        File file = new File("D:\\3.xls");
//        try  {
//            FileOutputStream fileOutputStream = new FileOutputStream(file);
//            resultWB.write(fileOutputStream);
//            fileOutputStream.close();
//        }catch (FileNotFoundException e){
//            e.printStackTrace();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }
        MergeAction mergeAction = new MergeAction();
        Main main = new Main();
        int status = main.readFiles("D:\\麻将.xlsx");
        File file = new File("D:\\1.xlsx");
        FileOutputStream fileOutputStream;
        try{
            fileOutputStream = new FileOutputStream(file);
            if(status==0){
                HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
                HSSFWorkbook bug = mergeAction.mergeHSSF(hssfWorkbook, "bug", "D:\\麻将.xlsx", "D:\\升级.xlsx");
                bug.write(fileOutputStream);
                fileOutputStream.close();
            }else {
                XSSFWorkbook xssfSheets = new XSSFWorkbook();
                XSSFWorkbook bug = mergeAction.mergeXSSF(xssfSheets, "bug", "D:\\麻将.xlsx", "D:\\升级.xlsx");
                bug.write(fileOutputStream);
                fileOutputStream.close();
            }
        }catch (IOException e){
            e.printStackTrace();
        }

    }
}
