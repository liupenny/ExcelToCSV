import com.xlscsv.converter.Converter;

import java.io.File;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by PennyLiu on 2016/12/8.
 */

public class TestMain {

    public static void main(String[] args) throws Exception{
        String basePath = System.getProperty("user.dir");
        String inputPath = basePath + "\\input\\";
        String outputPath = basePath + "\\output\\";
        String filename = inputPath + "test2.xls";
        //Converter.xls2csv(filename,outputPath);
        //Converter.xlsx2csv(filename,outputPath);
        Converter.csv2xls(inputPath,outputPath);
        //Converter.csv2xlsx(inputPath,outputPath);
    }

}
