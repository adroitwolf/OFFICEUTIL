package com.adroitwolf.entity;

import com.adroitwolf.anno.Excel;
import com.adroitwolf.anno.ExcelEntity;
import com.adroitwolf.enums.ExcelDataEnum;
import com.adroitwolf.enums.ExcelWorkEnum;
import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * <pre>Worker</pre>
 *
 * @author adroitwolf 2021年03月16日 10:28
 */
@Excel(height = 6,type = ExcelWorkEnum.XSSF,sheetName = "员工",path = "d://code",rowHeight = 2,rowNum = 2)
@Data
@AllArgsConstructor
public class Worker {

    @ExcelEntity(type = ExcelDataEnum.NUM,name = "编号")
    private Integer id;

    @ExcelEntity(mergeVertical = true,cellWidth = 2)
    private String name;

    @ExcelEntity(type = ExcelDataEnum.NUM)
    private Integer age;

    @ExcelEntity(name="地址",mergeVertical = true)
    private String address;
}
