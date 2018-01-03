package com.xlscsv.converter;

import com.opencsv.CSVReader;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;


class CsvToXlsx {
    CsvToXlsx(){

    }

    void process(String inputFileName, String outputFilePath, String sheetName) throws IOException{
        int i = 0;
        String s[];
        XSSFWorkbook wb;
        String outputFileName = outputFilePath + ".xlsx";
        File outputFile = new File(outputFileName);


        if(outputFile.exists()){
            wb = new XSSFWorkbook(new FileInputStream(outputFile));
        }else{
            wb = new XSSFWorkbook();
        }

        XSSFSheet sh = wb.createSheet(sheetName);
        XSSFRow row;
        XSSFCell cell;

        try{
            CSVReader reader = new CSVReader(new FileReader(inputFileName));
            while((s = reader.readNext()) !=null){
                row= sh.createRow(i);
                for(int j=0;j<s.length;j++){
                    cell= row.createCell(j);
                    cell.setCellValue(s[j]);
                }
                i+=1;
            }
        }
        catch(FileNotFoundException e){
            System.out.println("FileNotFound!");}
        catch (IOException e){
            e.printStackTrace();
        }

        FileOutputStream fout= new FileOutputStream(outputFileName);
        wb.write(fout);
        fout.close();
        wb.close();
    }
}
