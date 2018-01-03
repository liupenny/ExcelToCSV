package com.xlscsv.converter;

import java.io.FileNotFoundException;
import java.io.PrintStream;

public interface OutputDispatcher {
    public PrintStream openStreamForSheet(String sheetName) throws FileNotFoundException;
    public void closeStreamForSheet(PrintStream stream, String sheetName);
}