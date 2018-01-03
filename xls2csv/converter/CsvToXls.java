package com.xlscsv.converter;
import com.opencsv.CSVReader;

import java.io.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
class CsvToXls {
    CsvToXls(){

    }

    void process(String inputFileName, String outputFilePath, String sheetName) throws IOException{
        int i=0;
        String s[];
        Workbook wb;
        String outputFileName = outputFilePath + ".xls";
        File outputFile = new File(outputFileName);
        if(outputFile.exists()){
            wb = new HSSFWorkbook(new FileInputStream(outputFile));
        }else{
            wb = new HSSFWorkbook();
        }
//        Sheet sh= wb.createSheet("Sheet1");
        Sheet sh = wb.createSheet(sheetName);
        Row row;
        Cell cell;

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
//        System.out.println("Ok!");
    }
}
