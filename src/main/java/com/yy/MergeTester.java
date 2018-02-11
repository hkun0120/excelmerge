package com.yy;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;

/**
 * k
 * Created by storm on 2017/3/10.
 */
public class MergeTester {
    public static void main(String[] args) {
        System.out.printf("input your directory where xls files located: ");
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
        String path = null;
        try {
            path = reader.readLine();
        } catch (IOException e) {
            e.printStackTrace();
        }
        mergeFiles(path);
    }

    private static void mergeFiles(String path) {
        File file = new File(path);
        String destFile = "";

        List<String> listNames = new ArrayList<String>();
        if (file.isDirectory()){
            if (path.charAt(path.length() - 1) != '/') {
                path = path + "/";
            }
            destFile = path+"dest.xls";
            String[] fileNames =file.list();
            for (String fileName:fileNames) {
                if (fileName.contains(".xls")) {
                    listNames.add(path + fileName);
                    System.out.println(fileName);
                }
            }

        }

        try {
            ExcelMergeUtil.merge(listNames,destFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
