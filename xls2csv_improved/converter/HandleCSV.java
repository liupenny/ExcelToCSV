package com.xlscsv.converter;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.nio.charset.Charset;

import com.opencsv.CSVWriter;

public class HandleCSV {
	
	
	public static void main(String[] args) {
		HandleCSV hcsv=new HandleCSV();
    	String inputFileName="C:\\Users\\wky\\Desktop\\123.csv";

    	hcsv.read(inputFileName);
	}
	public void read(String inputFileName){
		try {    
			
	    	//String inputFileName="C:\\Users\\wky\\Desktop\\1.csv";

            BufferedReader reader = new BufferedReader(new FileReader(inputFileName));//换成你的文件名   
            reader.readLine();//第一行信息，为标题信息，不用,如果需要，注释掉   
            String line = null;    
            while((line=reader.readLine())!=null){    
                String item[] = line.split(",");//CSV格式文件为逗号分隔符文件，这里根据逗号切分   
                    
                String last = item[item.length-1];//这就是你要的数据了   
                System.out.println(last);    
            }    
        } catch (Exception e) {    
            e.printStackTrace();    
        }    
	}
	
//	public static void write(){
//
//        String filePath = "/Users/dddd/test.csv";
//
//        try {
//            // 创建CSV写对象
//            CSVWriter csvWriter = new CSVWriter(filePath,',', Charset.forName("utf-8"));
//            //CsvWriter csvWriter = new CsvWriter(filePath);
//
//            // 写表头
//            String[] headers = {"编号","姓名","年龄"};
//            String[] content = {"12365","张山","34"};
//            csvWriter.writeRecord(headers);
//            csvWriter.writeRecord(content);
//            csvWriter.close();
//
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }
	
}
