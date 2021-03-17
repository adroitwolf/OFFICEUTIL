package com.adroitwolf.entity;

import com.adroitwolf.enums.ExcelDataEnum;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

/**
 * <pre>ExportEntityList</pre>
 *
 * @author adroitwolf 2021年03月16日 15:29
 */
@Data
@NoArgsConstructor
public class ExportEntity {

    private int index;

    private Field field;

    // 字段的get方法
    private Method getMethod;

    private String name;

    private CellStyleEntity cellStyleEntity;

    private ExcelDataEnum type;

    private boolean mergeVertical;

    private int startNum = -1;

}
