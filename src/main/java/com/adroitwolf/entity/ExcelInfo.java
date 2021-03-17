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
    private String tableName;

    private ExcelWorkEnum type;

    private String sheetName;

    private boolean isHideTableHeader;

    private RowStyle rowStyle;

    private String path;
}
