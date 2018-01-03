package com.xlscsv.converter;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;

import java.io.File;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Converter {

    public static void xls2csv(String filename, String outputPath) throws Exception{
        File inputFile = new File(filename);
        String fileNameNoEx = getFileNameNoEx(inputFile.getName());
        OutputDispatcher dispatcher;
        dispatcher = new PerSheetOutputDispatcher(outputPath + fileNameNoEx);
        XlsToCsv xls2csv = new XlsToCsv(filename, dispatcher, -1);
        xls2csv.process();
    }

    //原来的
//    public static void xlsx2csv(String filename, String outputPath) throws Exception{
//        File inputFile = new File(filename);
//        String fileNameNoEx = getFileNameNoEx(inputFile.getName());
//        OutputDispatcher dispatcher;
//        dispatcher = new PerSheetOutputDispatcher(outputPath + fileNameNoEx);
//        OPCPackage p = OPCPackage.open(inputFile.getPath(), PackageAccess.READ);
//        XlsxToCsv xlsx2csv = new XlsxToCsv(p, dispatcher, -1);
//        xlsx2csv.process();
//    }

    public static void xlsx2csv(String filename, String outputPath) throws Exception{
        File inputFile = new File(filename);
        String fileNameNoEx = getFileNameNoEx(inputFile.getName());
        OutputDispatcher dispatcher;
        dispatcher = new PerSheetOutputDispatcher(outputPath + fileNameNoEx);
        OPCPackage p = OPCPackage.open(inputFile.getPath(), PackageAccess.READ);
        XLSX2CSV xls2csv = new XLSX2CSV(p, dispatcher, -1);
        xls2csv.process();
    }


    public static void csv2xls(String inputPath, String outputPath) throws Exception{
        File inputDir = new File(inputPath);
        File[] inputFiles = inputDir.listFiles();
        String pattern = "(.+)_(.+).csv";
        Pattern r = Pattern.compile(pattern);
        String fileName="";
        String sheetName="";

        for(File tmpFile:inputFiles) {
            Matcher m = r.matcher(tmpFile.getName());
            if (m.find()) {
                fileName = m.group(1);
                sheetName = m.group(2);
                new CsvToXls().process(inputPath + tmpFile.getName(), outputPath + fileName, sheetName);
            } else {
                fileName = Converter.getFileNameNoEx(tmpFile.getName());
                sheetName = "Sheet1";
                new CsvToXls().process(inputPath + tmpFile.getName(), outputPath + fileName, sheetName);
            }
        }
    }

    public static void csv2xlsx(String inputPath, String outputPath) throws Exception{
        File inputDir = new File(inputPath);
        File[] inputFiles = inputDir.listFiles();
        String pattern = "(.+)_(.+).csv";
        Pattern r = Pattern.compile(pattern);
        String fileName="";
        String sheetName="";

        for(File tmpFile:inputFiles) {
            Matcher m = r.matcher(tmpFile.getName());
            if (m.find()) {
                fileName = m.group(1);
                sheetName = m.group(2);
                new CsvToXlsx().process(inputPath + tmpFile.getName(), outputPath + fileName, sheetName);
            }
            else{
                fileName = Converter.getFileNameNoEx(tmpFile.getName());
                sheetName = "Sheet1";
                new CsvToXlsx().process(inputPath + tmpFile.getName(), outputPath + fileName, sheetName);
            }
        }
    }


    public static String getFileNameNoEx(String filename) {
        if ((filename != null) && (filename.length() > 0)) {
            int dot = filename.lastIndexOf('.');
            if ((dot >-1) && (dot < (filename.length()))) {
                return filename.substring(0, dot);
            }
        }
        return filename;
    }

}