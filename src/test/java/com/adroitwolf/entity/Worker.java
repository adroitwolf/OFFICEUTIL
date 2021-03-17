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
@Excel(height = 6,type = ExcelWorkEnum.XSSF,sheetName = "员工",path = "d://code",isHideTableFiled = false)
@Data
@AllArgsConstructor
public class Worker {

    @ExcelEntity(type = ExcelDataEnum.NUM,name = "编号")
    private Integer id;

    @ExcelEntity
    private String name;

    @ExcelEntity(type = ExcelDataEnum.NUM)
    private Integer age;


    private String address;
}
