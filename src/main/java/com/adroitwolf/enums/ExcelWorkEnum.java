package com.adroitwolf.enums;

/**
 * <pre>ExcelWorkEnum</pre>
 *
 * @author adroitwolf 2021年03月16日 09:01
 */
public enum ExcelWorkEnum {
    HSSF(".xls",65535),XSSF(".xlsx",65535),SXSSF(".xlsx",65535);

    private String suffix;

    private int size;


    ExcelWorkEnum(String suffix,int size) {
        this.suffix = suffix;
        this.size = size;
    }

    public String getSuffix(){
        return suffix;
    }

    public int getSize() {
        return size;
    }
}
