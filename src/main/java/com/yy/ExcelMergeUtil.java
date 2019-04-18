package com.yy;

import org.apache.poi.hssf.model.Sheet;
import org.apache.poi.hssf.model.Workbook;
import org.apache.poi.hssf.record.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.util.*;

/**
 * Created by storm on 2017/3/10.
 */
public class ExcelMergeUtil {
    public static void merge(List<String> sourceFiles, String destFile)throws Exception{
        InputStream[] inputs = new InputStream[sourceFiles.size()];
        for(int i=0; i<sourceFiles.size(); i++) {
            inputs[i] = new FileInputStream(sourceFiles.get(i));
        }

        OutputStream out = new FileOutputStream(destFile);

        merge(inputs, out);
    }

    public static void merge(InputStream[] inputs, OutputStream out)throws Exception{
        Map map = null;
        try{
            if(inputs == null || inputs.length <= 1) {
                throw new IllegalArgumentException("没有传入输入流数组或只有一个输入流！");
            }
            System.out.println("需要合并的文件数为：" + inputs.length);

            //第一个文档
            List<Record> rootRecords = getRecords(inputs[0]);
            Workbook workbook = Workbook.createWorkbook(rootRecords);
            List<Sheet> sheets = getSheets(workbook, rootRecords);
            if(sheets == null || sheets.size() == 0) {
                throw new IllegalArgumentException("第一个文档格式错误，必须至少有一个sheet！");
            }

            //以第一个文档的最后一个sheet为根，以后的数据都追加在这个sheet后面
            Sheet rootSheet = sheets.get(sheets.size() - 1);
            int rootRows = getRows(rootRecords); //记录第一篇文档的行数，以后的行数在此基础上增加
            rootSheet.setLoc(rootSheet.getDimsLoc());
            map = new HashMap(1000);

            for(int i = 1; i < inputs.length; i++){ //从第二篇开始遍历
                List<Record> records = getRecords(inputs[i]);
                int rowsOfCurXls = 0;
                //遍历当前文档的每一个record
                for(Iterator it = records.iterator(); it.hasNext();){
                    Record record = (Record) it.next();
                    if(record.getSid() == RowRecord.sid){ //如果是RowRecord
                        RowRecord rowRecord = (RowRecord) record;
                        rowRecord.setRowNumber(rootRows + rowRecord.getRowNumber()); //调整行号
                        rootSheet.addRow(rowRecord); //追加Row
                        rowsOfCurXls++; //记录当前文档的行数
                    }
                    //SST记录，SST保存xls文件中唯一的String，各个String都是对应着SST记录的索引
                    else if (record.getSid() == SSTRecord.sid){
                        SSTRecord sstRecord = (SSTRecord) record;
                        for (int j = 0; j < sstRecord.getNumUniqueStrings(); j++) {
                            int index = workbook.addSSTString(sstRecord.getString(j));
                            //记录原来的索引和现在的索引的对应关系
                            map.put(Integer.valueOf(j), Integer.valueOf(index));
                        }
                    }
                    else if (record.getSid() == LabelSSTRecord.sid){
                        LabelSSTRecord label = (LabelSSTRecord) record;
                        //调整SST索引的对应关系
                        label.setSSTIndex( ((Integer)map.get(Integer.valueOf(label.getSSTIndex()))).intValue() );
                    }

                    //追加ValueCell
                    if(record instanceof CellValueRecordInterface){
                        CellValueRecordInterface cell = (CellValueRecordInterface) record;
                        int cellRow = cell.getRow() + rootRows;
                        cell.setRow(cellRow);
                        rootSheet.addValueRecord(cellRow, cell);
                    }
                }
                rootRows += rowsOfCurXls;
            }

            byte[] data = getBytes(workbook, sheets.toArray(new Sheet[0]));

            write(out, data);

            System.out.println("合并完成");
        }finally{
            if(map!=null){
                map.clear();
                map = null;
            }
        }
    }

    static void write(OutputStream out, byte[] data)throws Exception{
        POIFSFileSystem fs = new POIFSFileSystem();
        try{
            fs.createDocument(new ByteArrayInputStream(data), "Workbook");
            fs.writeFilesystem(out);
            out.flush();
        }finally{
            try{
                out.close();
            }catch(IOException e){
                e.printStackTrace();
            }
        }
    }

    /**
     * 获取Sheet列表
     */
    static List<Sheet> getSheets(Workbook workbook, List<Record> records)throws Exception{
        int recOffset = workbook.getNumRecords();
        int sheetNum = 0;

        convertLabelRecords(records, recOffset, workbook);

        List<Sheet> sheets = new ArrayList<Sheet>();
        while(recOffset < records.size()){
            Sheet sheet = Sheet.createSheet(records, sheetNum++, recOffset);
            recOffset = sheet.getEofLoc() + 1;
            if(recOffset == 1) break;
            sheets.add(sheet);
        }
        return sheets;
    }

    /**
     * 取得一个sheet中数据的行数
     */
    static int getRows(List<Record> records) {
        int row = 0;
        for(Iterator it = records.iterator(); it.hasNext();){
            Record record = (Record) it.next();
            if(record.getSid() == DimensionsRecord.sid){
                DimensionsRecord dr = (DimensionsRecord)record;
                row = dr.getLastRow();
                break;
            }
        }
        return row;
    }

    /**
     * 获取Excel文档的记录集
     */
    public static List<Record> getRecords(InputStream input) {
        try{
            POIFSFileSystem poifs = new POIFSFileSystem(input);
            InputStream stream = poifs.getRoot().createDocumentInputStream("Workbook");
            return RecordFactory.createRecords(stream);
        }catch(IOException e){
            System.out.println("ExcelMergeUtil.getRecords: " + e.toString());
            e.printStackTrace();
        }
        return Collections.EMPTY_LIST;
    }

    static void convertLabelRecords(List<Record> records, int offset, Workbook workbook)throws Exception{
        for(int k = offset; k < records.size(); k++){
            Record rec = (Record) records.get(k);

            if (rec.getSid() == LabelRecord.sid) {
                LabelRecord oldrec = (LabelRecord) rec;

                records.remove(k);
                int stringid = workbook.addSSTString(new UnicodeString(oldrec.getValue()));

                LabelSSTRecord newrec = new LabelSSTRecord();
                newrec.setRow(oldrec.getRow());
                newrec.setColumn(oldrec.getColumn());
                newrec.setXFIndex(oldrec.getXFIndex());
                newrec.setSSTIndex(stringid);
                records.add(k, newrec);
            }
        }
    }

    static byte[] getBytes(Workbook workbook, Sheet[] sheets) {
        int nSheets = sheets.length;

        for(int i = 0; i < nSheets; i++){
            sheets[i].preSerialize();
        }

        int totalsize = workbook.getSize();

        int[] estimatedSheetSizes = new int[nSheets];
        for(int k = 0; k < nSheets; k++){
            workbook.setSheetBof(k, totalsize);
            int sheetSize = sheets[k].getSize();
            estimatedSheetSizes[k] = sheetSize;
            totalsize += sheetSize;
        }

        byte[] retval = new byte[totalsize];
        int pos = workbook.serialize(0, retval);

        for(int k = 0; k < nSheets; k++){
            int serializedSize = sheets[k].serialize(pos, retval);
            if(serializedSize != estimatedSheetSizes[k]){
                throw new IllegalStateException("Actual serialized sheet size (" + serializedSize
                        + ") differs from pre-calculated size (" + estimatedSheetSizes[k] + ") for sheet (" + k
                        + ")");
            }
            pos += serializedSize;
        }
        return retval;
    }

}
