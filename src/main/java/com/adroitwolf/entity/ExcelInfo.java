package com.adroitwolf.entity;

import com.adroitwolf.enums.ExcelWorkEnum;
import lombok.Data;

/**
 * <pre>ExcelInfo</pre>
 *
 * @author adroitwolf 2021年03月16日 11:17
 */
@Data
public class ExcelInfo {

    public ExcelInfo() {

    }

    public ExcelInfo(String tableName, ExcelWorkEnum type, String sheetName, boolean isHideTableHeader, RowStyle rowStyle, String path, Integer headerHeight,Integer dataRowHeight, Integer startRow) {
        this.tableName = tableName;
        this.type = type;
        this.sheetName = sheetName;
        this.isHideTableHeader = isHideTableHeader;
        this.rowStyle = rowStyle;
        this.path = path;
        this.headerHeight = headerHeight;
        this.startRow = startRow;
        this.dataRowHeight = dataRowHeight;
    }


    private String tableName;

    private ExcelWorkEnum type;

    private String sheetName;

    private boolean isHideTableHeader;

    private RowStyle rowStyle;

    private String path;

    private Integer headerHeight;

    private Integer dataRowHeight;

    private Integer startRow; // 起始行号

    private int size;  //一共多少个数据
}
