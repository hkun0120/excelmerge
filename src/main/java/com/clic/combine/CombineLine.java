package com.clic.combine;

import com.yy.ExcelMergeUtil;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author chinalife
 * @title: CombineLine
 * @projectName excelmerge
 * @description: 操作excel中的A\B两列，把A列的数据处理后当作成员变量注释，B列做驼峰处理当作字段名，把A列数据逐一较差放在B列
 * 上面，形成java类文件的样式
 * @date 2019/4/1718:00
 */
public class CombineLine {
    private static String SUCESS="sucess";
    private static String FAILD="faild";
    public String ProcessALine(String value){
        return "/** "+value+" **/";
    }


    /**
     * @description: 把带下划线的值转换为驼峰规则的值
     * @author
     * @date 2019/4/17 18:23
     */
    public String ProcessBLine(String value){
        String hump = lineToHump(value);
        return "private String "+hump+";";
    }

    /** 下划线转驼峰 */
    public String lineToHump(String str) {
        Pattern linePattern = Pattern.compile("_(\\w)");
        str = str.toLowerCase();
        Matcher matcher = linePattern.matcher(str);
        StringBuffer sb = new StringBuffer();
        while (matcher.find()) {
            matcher.appendReplacement(sb, matcher.group(1).toUpperCase());
        }
        matcher.appendTail(sb);
        return sb.toString();
    }

    /**
     * @description: 文件转化为流
     * @author hk
     * @date 2019/4/17 18:41
     */
    private String ReadWriteXls(String srcPath,String destPath){
        File file = new File(srcPath);
        try {
            InputStream inputStream = new FileInputStream(file);
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook(inputStream);
            for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++){

            }
            HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
            for (int rowNum = 0; rowNum<=hssfSheet.getLastRowNum();rowNum++){
                HSSFRow hssfRow = hssfSheet.getRow(rowNum);
                if(hssfRow != null){
                    HSSFCell hssfCellA = hssfRow.getCell(0);
                    HSSFCell hssfCellB = hssfRow.getCell(1);
                    System.out.println(ProcessALine(getValue(hssfCellA)));
                    System.out.println(ProcessBLine(getValue(hssfCellB)));
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            return FAILD;
        } catch (IOException e) {
            e.printStackTrace();
            return FAILD;
        }
        return SUCESS;
    }


    private String getValue(HSSFCell hssfRow) {
       if (hssfRow.getCellType() == hssfRow.CELL_TYPE_BOOLEAN) {
         return String.valueOf(hssfRow.getBooleanCellValue());
        } else if (hssfRow.getCellType() == hssfRow.CELL_TYPE_NUMERIC) {
         return String.valueOf(hssfRow.getNumericCellValue());
          } else {
            return String.valueOf(hssfRow.getStringCellValue());
          }
    }
    public static void main(String[] args) {
        CombineLine combineLine = new CombineLine();
//        System.out.println(new CombineLine().lineToHump("INSUR_DUR_UNIT"));
        combineLine.ReadWriteXls("D:\\3document\\1.xls","");
    }
}
