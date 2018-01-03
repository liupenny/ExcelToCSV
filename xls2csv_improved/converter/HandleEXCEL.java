package com.xlscsv.converter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HandleEXCEL {
	
	static ArrayList<String> ArrayString = new ArrayList<String>();
	
	public static void main(String[] args) {
		HandleEXCEL handle = new HandleEXCEL();
//		handle.deleteFiles("C:\\Users\\wky\\Desktop\\单机版邮件.xls");
//		String path="C:\\Users\\wky\\Desktop\\邮件.xls";
//		String path="C:\\Users\\wky\\Desktop\\邮件.xls";
//		String path1="C:\\Users\\wky\\Desktop\\邮件temp.xls";
		String path="/home/hadoop/桌面/邮件.xls";
		
//		String path1="C:\\Users\\wky\\Desktop\\单机版邮件temp.xls";
		String path1="/home/hadoop/桌面/邮件temp.xls";
//		System.out.println(new File(path).getAbsolutePath());
//		System.out.println(new File(path).getName());
//		System.out.println(new File(path).getParentFile());
		handle.AddNumExcel(path);
//		handle.RemoveNumExcel(path1);
	}
	
	public void AddNumExcel(String path){
		String fileType = path.substring(path.lastIndexOf(".") + 1, path.length());
 		// 创建工作文档对象
 		Workbook wb = null;
 		String outpath="";
 		if (fileType.equals("xls")) {
 			wb = new HSSFWorkbook();
 			outpath=path.substring(0, path.length()-4)+"temp.xls";
 		} else if (fileType.equals("xlsx")) {
 			wb = new XSSFWorkbook();
 			outpath=path.substring(0, path.length()-5)+"temp.xlsx";
 		} else {
 			System.out.println("您的文档格式不正确！");
 		}
 		
    	if(path.endsWith(".xlsx")){

			try {
	    		File file = new File(path);
	            
	            InputStream in = new FileInputStream(file);
	            XSSFWorkbook excel= new XSSFWorkbook(in);//得到整个excel对象
						
	            int sheets = excel.getNumberOfSheets();		//获取整个excel有多少个sheet
	            XSSFRow row1;
	    		HandleEXCEL h=new HandleEXCEL();
	            for(int i = 0 ; i < sheets ; i++ ){		//遍历每一个sheet
	            	
	                XSSFSheet sheet = excel.getSheetAt(i);
	                
	                ArrayString.add(sheet.getSheetName());
	                
	                if(sheet.getLastRowNum()==0){
	                	 Sheet sheet1 = (Sheet) wb.createSheet(sheet.getSheetName());	//创建sheet对象
	                    continue;
	                }
//	                int number=sheet.getLastRowNum()+1;	//获取excel文件中数据的行数  包括了两行之间空的一行的数目
	                int number=h.getRealRowNumxlsx(sheet);
	                Sheet sheet1 = (Sheet) wb.createSheet(sheet.getSheetName());	//创建sheet对象
	                
	                for( int rowNum = 0 ; rowNum < number; rowNum++ ){		//遍历一个sheet中的每一行
	                    row1 = sheet.getRow(rowNum);
	                    
	                    if(h.isXSSNull(row1)){	//中间某一行为空的情况，由于没有数据，跳出循环
	                    	Row row = (Row) sheet1.createRow(rowNum);	//写入一行空的数据
	                    	Cell cell = row.createCell(0);
	                    	cell.setCellValue(""+(rowNum+1));
	                        continue;
	                    }
	                    
	                    int columnNum=row1.getLastCellNum();	//获取一行中有多少列（修改后可能不止6列）
	                    
	                    Row row = (Row) sheet1.createRow(rowNum);
	                   
	                    for( int col = 0 ; col < columnNum+1 ; col++ ){	//对每一行中的列进行遍历
	                    	
	                         Cell cell = row.createCell(col);
	                         if(col==0){
	                        	 cell.setCellValue(""+(rowNum+1));
	                         }else{
	                        	 XSSFCell cell1 = row1.getCell(col-1);
	                        	 if(getCellValue(cell1)==null)
	                        		 cell.setCellValue("");
		                         else
		                        	 cell.setCellValue(getCellValue(cell1));
	                        	 }
	                    }
	    	}
	                }
	 			OutputStream stream = new FileOutputStream(outpath);
	 			wb.write(stream);
	 			stream.close();
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
    	}else{
			try {
	    		File file = new File(path);
	            
	            InputStream in = new FileInputStream(file);
	            HSSFWorkbook excel= new HSSFWorkbook(in);//得到整个excel对象
						
	            int sheets = excel.getNumberOfSheets();		//获取整个excel有多少个sheet
	            HSSFRow row1;
	            HandleEXCEL h=new HandleEXCEL();
	            for(int i = 0 ; i < sheets ; i++ ){		//遍历每一个sheet
	                HSSFSheet sheet = excel.getSheetAt(i);
	                ArrayString.add(sheet.getSheetName());
	                if(sheet.getLastRowNum()==0||sheet==null){
		                Sheet sheet1 = (Sheet) wb.createSheet(sheet.getSheetName());	//创建sheet对象
		                //System.out.println(sheet.getSheetName()+"行数为"+(sheet.getLastRowNum()+1));
		                continue;
	                }
	               int number=h.getRealRowNumxls(sheet);
	              
	                Sheet sheet1 = (Sheet) wb.createSheet(sheet.getSheetName());	//创建sheet对象
	                
	                for( int rowNum = 0 ; rowNum < number; rowNum++ ){		//遍历一个sheet中的每一行
	                    row1 = sheet.getRow(rowNum);
//	                    System.out.println("遍历行数"+(rowNum+1));
	                    if(h.isHSSNull(row1)){	//中间某一行为空的情况，由于没有数据，跳出循环
	                    	Row row = (Row) sheet1.createRow(rowNum);	//写入一行空的数据
	                    	Cell cell = row.createCell(0);
//	                    	System.out.println(sheet.getSheetName()+"空行情况"+(rowNum+1));
	                    	cell.setCellValue(""+(rowNum+1));
	                        continue;
//	                    }else{
//	                    	System.out.println("非空行"+(rowNum+1)+"    "+row1.getPhysicalNumberOfCells());
	                    }
	                    
	                    int columnNum=row1.getLastCellNum();	//获取一行中有多少列（修改后可能不止6列）
	                    
	                    Row row = (Row) sheet1.createRow(rowNum);
	                   
	                    for( int col = 0 ; col < columnNum+1 ; col++ ){	//对每一行中的列进行遍历
	                    	
	                         Cell cell = row.createCell(col);
	                         if(col==0){
	                        	 cell.setCellValue(""+(rowNum+1));
	                         }else{
	                        	 HSSFCell cell1 = row1.getCell(col-1);
	                        	 if(getCellValue(cell1)==null)
	                        		 cell.setCellValue("");
		                         else
		                        	 cell.setCellValue(getCellValue(cell1));
	                        	 }
	                    }
	                        	 
	    	}
	                }
	 			OutputStream stream = new FileOutputStream(outpath);
	 			wb.write(stream);
	 			stream.close();
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
          
    	
    	}
	}

	//针对xls文件，获取一个sheet表中的实际行数
			public int getRealRowNumxls(HSSFSheet sheet){
				
				int num=sheet.getLastRowNum();
				HandleEXCEL h=new HandleEXCEL();
				HSSFRow row1 = sheet.getRow(num);
				if(!h.isHSSNull(row1)){
					return num+1;
				}else{
					num--;
					row1 = sheet.getRow(num);
					while(h.isHSSNull(row1)){
						num--;
						row1 = sheet.getRow(num);
					}
				}
				return num+1;
			}
			
			//针对xlsx文件，获取一个sheet表中的实际行数
			public int getRealRowNumxlsx(XSSFSheet sheet){
				
				int num=sheet.getLastRowNum();
				HandleEXCEL h=new HandleEXCEL();
				XSSFRow row1 = sheet.getRow(num);
				if(!h.isXSSNull(row1)){
					return num+1;
				}else{
					num--;
					row1 = sheet.getRow(num);
					while(h.isXSSNull(row1)){
						num--;
						row1 = sheet.getRow(num);
					}
				}
				return num+1;
			}
		//判断xlsx文件中一行是否为“空”
		public boolean isXSSNull(XSSFRow row1){
			
			if(row1==null)
				return true;
			boolean b=true;
			int num=row1.getLastCellNum();
			for(int i=0;i<num;i++){
				if(row1.getCell(i)==null||row1.getCell(i).toString().equals("")){
					b=true;
				}
				else{
					b=false;
					break;
				}
			}
			return b;
		}
		//判断xls文件中一行是否为“空”
		public boolean isHSSNull(HSSFRow row1){
			if(row1==null)
				return true;
			boolean b=true;
			int num=row1.getLastCellNum();
			for(int i=0;i<num;i++){
				if(row1.getCell(i)==null||row1.getCell(i).toString().equals("")){
					b=true;
				}
				else{
					b=false;
					break;
				}
			}
			return b;
		}
	
//	 public static String getCellValue(Cell cell) {
//	        if (cell == null) {
//	            return null;
//	        }
//	        switch (cell.getCellType()) {
//	            case Cell.CELL_TYPE_STRING:
//	                return cell.getRichStringCellValue().getString().trim();
//	            case Cell.CELL_TYPE_NUMERIC:
//	                if (DateUtil.isCellDateFormatted(cell)) {
//	                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");//非线程安全
//	                    return sdf.format(cell.getDateCellValue());
//	                } else {
//	                	if(cell.getNumericCellValue()==(int)cell.getNumericCellValue()){
//	                		return String.valueOf((int)cell.getNumericCellValue());
//	                	}else if(cell.getNumericCellValue()>0&&cell.getNumericCellValue()<1){
//	                		System.out.println(cell.getNumericCellValue());
//	                		return ""+cell.getNumericCellValue();
//	                	}else{
//	                		return String.valueOf(cell.getNumericCellValue());
//	                	}
//	                }
//	            case Cell.CELL_TYPE_BOOLEAN:
//	                return String.valueOf(cell.getBooleanCellValue());
//	            case Cell.CELL_TYPE_FORMULA:
//	                return cell.getCellFormula();
//	            default:
//	                return null;
//	        }
//	    }
		public static String getCellValue(Cell cell) {
	        if (cell == null) {
	            return null;
	        }

	        switch (cell.getCellType()) {
	            case Cell.CELL_TYPE_STRING:
	                return cell.getRichStringCellValue().getString().trim();
	            case Cell.CELL_TYPE_NUMERIC:
	                if (DateUtil.isCellDateFormatted(cell)) {
	                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");//非线程安全
	                    return sdf.format(cell.getDateCellValue());
	                } else {
	                	if(cell.getNumericCellValue()==(int)cell.getNumericCellValue()){
	                		return String.valueOf((int)cell.getNumericCellValue());
	                	}else if(cell.getNumericCellValue()>0&&cell.getNumericCellValue()<1){
	                		return ""+cell.getNumericCellValue();
	                	}else{
	                		return String.valueOf(cell.getNumericCellValue());
	                	}
	                }
	            case Cell.CELL_TYPE_BOOLEAN:
	                return String.valueOf(cell.getBooleanCellValue());
	            case Cell.CELL_TYPE_FORMULA:
	                return cell.getCellFormula();
	            default:
	                return null;
	        }
	    }
	    //处理公式
	    public static String getCellValueFormula(Cell cell, FormulaEvaluator formulaEvaluator) {
	        if (cell == null || formulaEvaluator == null) {
	            return null;
	        }

	        if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
	            return String.valueOf(formulaEvaluator.evaluate(cell).getNumberValue());
	        }
	        return getCellValue(cell);
	    }
	
	public void RemoveNumExcel(String path){

		String fileType = path.substring(path.lastIndexOf(".") + 1, path.length());
 		// 创建工作文档对象
 		Workbook wb = null;
 		String outpath="";
 		if (fileType.equals("xls")) {
 			wb = new HSSFWorkbook();
 			outpath=path.substring(0, path.length()-8)+".xls";
 		} else if (fileType.equals("xlsx")) {
 			wb = new XSSFWorkbook();
 			outpath=path.substring(0, path.length()-9)+".xlsx";
 		} else {
 			System.out.println("您的文档格式不正确！");
 		}
 		
    	if(path.endsWith(".xlsx")){
			try {
	    		File file = new File(path);
	            
	            InputStream in = new FileInputStream(file);
	            XSSFWorkbook excel= new XSSFWorkbook(in);//得到整个excel对象
						
	            int sheets = excel.getNumberOfSheets();		//获取整个excel有多少个sheet
	            XSSFRow row1;
	    		HandleEXCEL h=new HandleEXCEL();
	            for(int i = 0 ; i < sheets ; i++ ){		//遍历每一个sheet
	            	
	            	 XSSFSheet sheet;
	            	if(ArrayString.size()!=0){
	            		sheet=excel.getSheet(ArrayString.get(i));
	            	}else{
	            		sheet = excel.getSheetAt(i);
	            	}
	                if(sheet.getLastRowNum()==0){
		                Sheet sheet1 = (Sheet) wb.createSheet(sheet.getSheetName());	//创建sheet对象
	                    continue;
	                }
//	                int number=sheet.getLastRowNum()+1;	//获取excel文件中数据的行数  包括了两行之间空的一行的数目
	                int number=h.getRealRowNumxlsx(sheet);
	                Sheet sheet1 = (Sheet) wb.createSheet(sheet.getSheetName());	//创建sheet对象
	                
	                for( int rowNum = 0 ; rowNum < number; rowNum++ ){		//遍历一个sheet中的每一行
	                    row1 = sheet.getRow(rowNum);
	                    
	                    int columnNum=row1.getLastCellNum();
	                    Row row = (Row) sheet1.createRow(rowNum);
	                    if(columnNum==1){
	                         Cell cell = row.createCell(0);
	                         cell.setCellValue("");
	                    }else{
	                    	for( int col = 0 ; col < columnNum-1 ; col++ ){	//对每一行中的列进行遍历
		                         Cell cell = row.createCell(col);
		                         cell.setCellValue(row1.getCell(col+1).toString());
	                    	}
	                    }
	                    
	                }
	            }
	 			OutputStream stream = new FileOutputStream(outpath);
	 			wb.write(stream);
	 			stream.close();
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
          ArrayString.clear();
    	}else{
			try {
	    		File file = new File(path);
	            
	            InputStream in = new FileInputStream(file);
	            HSSFWorkbook excel= new HSSFWorkbook(in);//得到整个excel对象
						
	            int sheets = excel.getNumberOfSheets();		//获取整个excel有多少个sheet
	            HSSFRow row1;
	    		HandleEXCEL h=new HandleEXCEL();
	            for(int i = 0 ; i < sheets ; i++ ){		//遍历每一个sheet
	            	
	                HSSFSheet sheet;
	                if(ArrayString.size()!=0){
	            		sheet=excel.getSheet(ArrayString.get(i));
	            	}else{
	            		sheet = excel.getSheetAt(i);
	            	}
	                if(sheet.getLastRowNum()==0){
		                Sheet sheet1 = (Sheet) wb.createSheet(sheet.getSheetName());	//创建sheet对象
	                    continue;
	                }
//	                int number=sheet.getLastRowNum()+1;	//获取excel文件中数据的行数  包括了两行之间空的一行的数目
	                int number =h.getRealRowNumxls(sheet);
	                Sheet sheet1 = (Sheet) wb.createSheet(sheet.getSheetName());	//创建sheet对象
	                
	                for( int rowNum = 0 ; rowNum < number; rowNum++ ){		//遍历一个sheet中的每一行
	                    row1 = sheet.getRow(rowNum);
	                    
	                    int columnNum=row1.getLastCellNum();	//获取一行中有多少列（修改后可能不止6列）
	                    
	                    Row row = (Row) sheet1.createRow(rowNum);
	                   
	                    if(columnNum==1){
	                    	 HSSFCell cell1 = row1.getCell(0);
	                         Cell cell = row.createCell(0);
	                         cell.setCellValue("");
	                    }else{
	                    	for( int col = 0 ; col < columnNum-1 ; col++ ){	//对每一行中的列进行遍历
		                    	 HSSFCell cell1 = row1.getCell(col);
		                         Cell cell = row.createCell(col);
		                         cell.setCellValue(row1.getCell(col+1).toString());
	                    	}
	                    }
	                }
	            }
	 			OutputStream stream = new FileOutputStream(outpath);
	 			wb.write(stream);
	 			stream.close();
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
    	ArrayString.clear();
    	}
	}
	
	public void deleteFiles(String path){
	       File file = new File(path);
	       //1級文件刪除
	       if(!file.isDirectory()){
	           file.delete();
	       }else if(file.isDirectory()){
	           //2級文件列表
	           String []filelist = file.list();
	           //獲取新的二級路徑
	           for(int j=0;j<filelist.length;j++){
	               File filessFile= new File(path+"\\"+filelist[j]);
	               if(!filessFile.isDirectory()){
	                   filessFile.delete();
	               }else if(filessFile.isDirectory()){
	                   //遞歸調用
	                   deleteFiles(path+"\\"+filelist[j]);
	               }
	           }
	           file.delete();
	       }
	    }
}