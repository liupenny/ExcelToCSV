package com.xlscsv.converter;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.PrintStream;

public class PerSheetOutputDispatcher implements OutputDispatcher {
    private String basePath;
    PerSheetOutputDispatcher(String basePath) {
        this.basePath = basePath;
    }

    public PrintStream openStreamForSheet(String sheetName) throws FileNotFoundException {
        return new PrintStream(new File(basePath + "_" + sheetName + ".csv"));
    }

    public void closeStreamForSheet(PrintStream stream, String sheetName) {
        stream.close();
    }
}