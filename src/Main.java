import com.pactera.merge.MergeAction;
import com.pactera.merge.MergePlus;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;

public class Main {
    public static void main(String[] args) {
        MergePlus mergePlus = new MergePlus();
        MergeAction mergeAction = new MergeAction();
        try{
            File targetFile = mergePlus.readFile("D:\\1.xlsx");
            File[] sourceFiles = mergePlus.readDir("D:\\周报");
            System.out.println(sourceFiles.length+"个文件");
            Workbook targetWorkbook = mergePlus.excelEndWith(sourceFiles[0],0);
            Workbook result = mergeAction.mergeLastSheet(targetWorkbook, sourceFiles);
            FileOutputStream out = new FileOutputStream(targetFile);
            result.write(out);
            out.close();
        }catch (Exception e){
            e.printStackTrace();
        }

    }
}
