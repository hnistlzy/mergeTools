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
    private File readFiles(String ...path){
        File file ;
        String str=path[0];
        if(str.contains(".xlsx")){
             file = new File("D:\\aaa.xlsx");
        }else{
            file = new File("D:\\aaa.xls");
        }
        return  file;

}
    public static void main(String[] args) {
        MergeAction mergeAction = new MergeAction();
        Main main = new Main();
        File resultFile = main.readFiles("D:\\1.xls");
        FileOutputStream fileOutputStream;
        try{
            fileOutputStream = new FileOutputStream(resultFile);
            if(!resultFile.getName().contains(".xlsx")){
                HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
                HSSFWorkbook bug = mergeAction.mergeHSSF(hssfWorkbook, "bug", "D:\\1.xls", "D:\\2.xls");
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
