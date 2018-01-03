package com.xlscsv.converter;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.InputStreamReader;
import java.util.Scanner;


public class HandleTXT{


    /**
     * @param args
     */
    public static void main(String[] args){
    	HandleTXT handle=new HandleTXT();
    	handle.AddNumTxt("/home/hadoop/桌面/mail.txt");
//    	handle.RemoveNumTxt("/home/hadoop/桌面/mailtemp.txt");
    }
    
    
    public void AddNumTxt(String pathname){
    	  
        try { // 防止文件建立或读取失败，用catch捕捉错误并打印，也可以throw  

            /* 读入TXT文件 */  
        	//pathtmp用于存放新建的加标号的临时文件
        	String pathtmp=pathname.substring(0,pathname.length()-4)+"temp.txt";
            File filename = new File(pathname); // 要读取以上路径的input。txt文件  
            InputStreamReader reader = new InputStreamReader(  
                    new FileInputStream(filename),"utf-8"); // 建立一个输入流对象reader  
            BufferedReader br = new BufferedReader(reader); // 建立一个对象，它把文件内容转成计算机能读懂的语言  
            String line = "";  
            String Line = "";  
            int i=0;
            while ((line=br.readLine())!=null) {  
            	i++;
            	line=i+"  "+line;
                Line=Line+line+"\r\n";
            }  
            /* 写入Txt文件 */  
            File writename = new File(pathtmp); // 相对路径，如果没有则要建立一个新的output。txt文件  
            writename.createNewFile(); // 创建新文件  
            BufferedWriter out = new BufferedWriter(new FileWriter(writename));  
            out.write(Line); // \r\n即为换行  
            out.flush(); // 把缓存区内容压入文件  
            out.close(); // 最后记得关闭文件  

        } catch (Exception e) {  
            e.printStackTrace();  
        }  
    
    }
     
    public void RemoveNumTxt(String pathname){
  	  
        try { // 防止文件建立或读取失败，用catch捕捉错误并打印，也可以throw  

            /* 读入TXT文件 */  
        	//pathtmp用于存放新建的加标号的临时文件
        	String pathtmp=pathname.substring(0,pathname.length()-4)+".txt";
            File filename = new File(pathname); // 要读取以上路径的input。txt文件  
            InputStreamReader reader = new InputStreamReader(  
                    new FileInputStream(filename),"utf-8"); // 建立一个输入流对象reader  
            BufferedReader br = new BufferedReader(reader); // 建立一个对象，它把文件内容转成计算机能读懂的语言  
            String line = "";  
            String Line = "";  
            int i=0;
            while ((line=br.readLine())!=null) {  
            	i++;
            	line=line.substring((line.split("  ")[0]+"  ").length());
                Line=Line+line+"\r\n";
            }  
            /* 写入Txt文件 */  
            File writename = new File(pathtmp); // 相对路径，如果没有则要建立一个新的output。txt文件  
            writename.createNewFile(); // 创建新文件  
            BufferedWriter out = new BufferedWriter(new FileWriter(writename));  
            out.write(Line); // \r\n即为换行  
            out.flush(); // 把缓存区内容压入文件  
            out.close(); // 最后记得关闭文件  

        } catch (Exception e) {  
            e.printStackTrace();  
        }  
    }
    
    //实现删除指定路径下的文件的功能
    //由于分布式是删除整个文件夹，不用再单独删除文件了
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