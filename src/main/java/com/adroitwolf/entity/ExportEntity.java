package com.adroitwolf.entity;

import com.adroitwolf.enums.ExcelDataEnum;
import lombok.Data;
import lombok.NoArgsConstructor;
import java.lang.reflect.Method;

/**
 * <pre>ExportEntityList</pre>
 *
 * @author adroitwolf 2021年03月16日 15:29
 */
@Data
@NoArgsConstructor
public class ExportEntity {

    private int index; // 最开始的列索引

    private int endNum; //最终的列索引

    private int colNum; // 占多少列

    private int startRowNum = -1;

    private Object beforeValue;

    // 字段的get方法
    private Method getMethod;

    private String name;

    private CellStyleEntity cellStyleEntity;

    private ExcelDataEnum type;

    private boolean mergeVertical;


}
