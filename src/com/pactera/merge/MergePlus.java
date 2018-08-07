package com.pactera.merge;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;

public class MergePlus {
    /**
     * 从指定路径读取一个文件，若该文件存在则抛出异常
     * 不存在则创建
     * @param targetPath 文件目录
     * @return 文件对象
     * @throws FileReadException 自定义异常
     */
    public File readFile(String targetPath) throws FileReadException, IOException {
        File targetFile = new File(targetPath);
        if(targetFile.exists()){
            throw new FileReadException("该目标文件已经存在！");
        }else if(targetFile.isDirectory()){
            throw  new FileReadException("这是一个目录！");
        }else{
            System.out.println("文件"+targetPath+"创建成功");
        }
        return targetFile;
    }

    /**
     * 读取一个目录，并判断用户输入是否合法
     * @param sourcePath 源文件目录
     * @return 该目录下所有文件的数组
     * @throws FileReadException 自定义异常
     */
    public File[] readDir(String sourcePath) throws FileReadException {
        File sourceDir = new File(sourcePath);
        File[] files = sourceDir.listFiles();
        if(!sourceDir.isDirectory()){
            throw new FileReadException("请输入一个文件目录");
        }else if (files==null ||files.length<=0){
            throw new FileReadException("该目录下没有任何文件！");
        }
        return files;
    }


    /**
     * 根据文件的后缀名来判断创建哪种对象
     * @param file 文件，
     * @param status 是否读入文件，>0 读入文件，==0不读入文件，只根据文件名返回wb
     * @return
     * @throws FileReadException
     * @throws IOException
     */
    public Workbook excelEndWith(File file,Integer status) throws FileReadException,IOException {
        Workbook workbook;
        InputStream in;
        if(status==null||status==0){
            if(file.getName().endsWith(".xlsx")){
                workbook=new XSSFWorkbook();
            }else if(file.getName().endsWith(".xls")){
                workbook=new HSSFWorkbook();
            }else{
                throw  new FileReadException("请输入一个excel格式的文件");
            }
        }else{
            in =new FileInputStream(file);
            if(file.getName().endsWith(".xlsx")){
                workbook=new XSSFWorkbook(in);
            }else if(file.getName().endsWith(".xls")){
                workbook=new HSSFWorkbook(in);
            }else{
                throw  new FileReadException("请输入一个excel格式的文件");
            }
        }

        return workbook;
    }




    /**
     * 复制文件的第i个sheet的第一行数据，到目标sheet
     * @param targetSheet 目标sheet
     * @param sourceFile 文件，
     * @param sheetAt 第几个sheet
     * @throws IOException io
     * @throws FileReadException 自定义异常
     * */
    public void copyFirstRow(Sheet targetSheet, File sourceFile,int sheetAt) throws IOException, FileReadException {
        Workbook sourceWorkbook = excelEndWith(sourceFile,1);

        Sheet sourceSheet = sourceWorkbook.getSheetAt(sheetAt);
        if(sourceSheet.getLastRowNum()>0){
            Row sourceRow = sourceSheet.getRow(0);
            int firstCellNum = sourceRow.getFirstCellNum();
            int lastCellNum = (int)sourceRow.getLastCellNum();
            Row targetRow = targetSheet.createRow(0);
            FormulaEvaluator formulaEvaluator = sourceWorkbook.getCreationHelper().createFormulaEvaluator();
            for(;firstCellNum<lastCellNum;firstCellNum++){
                Cell targetCell = targetRow.createCell(firstCellNum);
                copyCell(sourceRow.getCell(firstCellNum),targetCell,formulaEvaluator);
            }
        }

    }

    /**
     *  复制整个sheet
     * @param sourceSheet 源sheet
     * @param targetSheet 目标sheet
     * @return 是否成功
     */
    public boolean copySheet(Sheet sourceSheet, Sheet targetSheet) {
        System.out.println("当前目标sheet的名字是："+targetSheet.getSheetName());
       int lastRowNum = targetSheet.getLastRowNum();
       System.out.println("目标sheet的最后一行是："+lastRowNum);
       //从第2行开始复制数据
       int firstNum = sourceSheet.getFirstRowNum()+1;
       int lastNum = sourceSheet.getLastRowNum();
       System.out.println("当前文件的首行是："+(firstNum)+"末行是:"+lastNum+"一共要读取"+(lastNum-firstNum+1)+"行数据");

       for(;firstNum<=lastNum;firstNum++){
           System.out.println("正在复制第"+firstNum+"行数据");
           //获取第一列 和最后一列的列号.
           int firstCellNum = (int)sourceSheet.getRow(firstNum).getFirstCellNum();
           int  lastCellNum = (int)sourceSheet.getRow(firstNum).getLastCellNum();
           //targetSheet新建一行
           Row targetRow = targetSheet.createRow(lastRowNum+1);
           for(;firstCellNum<=lastCellNum;firstCellNum++){
               Cell targetRowCell = targetRow.createCell(firstCellNum);
               Cell sourceCell = sourceSheet.getRow(firstNum).getCell(firstCellNum);
               FormulaEvaluator formulaEvaluator = sourceSheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
               if(sourceCell!=null){
                   copyCell(sourceCell,targetRowCell,formulaEvaluator);
               }else{
                   targetRowCell.setCellValue(" ");
               }
           }
           lastRowNum++;
       }


       return false;
   }

    /**
     *  复制cell中的值到targetCell
     * @param sourceCell 源cell
     * @param targetCell 目标cell
     * @param formulaEvaluator xx
     * @return 是否复制成功
     */
    public boolean copyCell(Cell sourceCell, Cell targetCell, FormulaEvaluator formulaEvaluator) {
       boolean bool=false;
        switch (sourceCell.getCellType()){
            case Cell.CELL_TYPE_BLANK:
                targetCell.setCellValue(" ");
                bool=true;
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                bool=true;
                break;
            case Cell.CELL_TYPE_ERROR:
                targetCell.setCellValue(sourceCell.getErrorCellValue());
                bool=true;
                break;
            case Cell.CELL_TYPE_STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                bool=true;
                break;
            case Cell.CELL_TYPE_NUMERIC:
                numberOfDate(targetCell,sourceCell.getCellStyle().getDataFormat(),sourceCell.getNumericCellValue(),"MM/dd","HH:mm");
                bool=true;
                break;
            case Cell.CELL_TYPE_FORMULA:
                CellValue evaluate = formulaEvaluator.evaluate(sourceCell);
                double numberValue = evaluate.getNumberValue();
                targetCell.setCellValue(numberValue);
                bool=true;
                break;
        }
        return bool;

    }

    /**
     *  格式化日期类型
     * @param targetCell  目标cell
     * @param dataFormat dataFormat
     * @param value sourceCell的值
     * @param dateFormat dateFormat
     * @param timeFormat timeFormat
     */
    private void numberOfDate(Cell targetCell, short dataFormat, double value, String dateFormat, String timeFormat) {
        SimpleDateFormat sdf ;
        if (dataFormat == 14 || dataFormat == 31 || dataFormat == 57 || dataFormat == 58
                || (176<=dataFormat && dataFormat<=178) || (182<=dataFormat && dataFormat<=196)
                || (210<=dataFormat &&dataFormat<=213) || (208==dataFormat ) ) { // 日期
            sdf = new SimpleDateFormat(dateFormat);
            Date date = DateUtil.getJavaDate(value);
            targetCell.setCellValue(sdf.format(date));
        } else if (dataFormat == 20 ||dataFormat == 32  || (200<=dataFormat && dataFormat<=209) ) { // 时间
            sdf = new SimpleDateFormat(timeFormat);
            Date date = DateUtil.getJavaDate(value);

            targetCell.setCellValue(sdf.format(date));
        }else{

            targetCell.setCellValue(value);
        }
    }


}
