package com.adroitwolf.enums;

/**
 * <pre>ExcelWorkEnum</pre>
 *
 * @author adroitwolf 2021年03月16日 09:01
 */
public enum ExcelWorkEnum {
    HSSF(".xls"),XSSF(".xlsx"),SXSSF(".xlsx");

    private String suffix;


    ExcelWorkEnum(String suffix) {
        this.suffix = suffix;
    }

    public String getSuffix(){
        return suffix;
    }
}
