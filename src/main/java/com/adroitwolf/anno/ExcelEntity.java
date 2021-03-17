package com.adroitwolf.anno;

import com.adroitwolf.enums.ExcelDataEnum;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * <pre>ExcelEntity</pre> 如果没有添加该注解，则默认不导入该属性
 *
 * @author adroitwolf 2021年03月16日 09:43
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelEntity {
    /**
     * 该单元格所占用的空间
     */
    public int width() default  5;

    /**
     * 该单元格的数据类型,默认为文本
     */
    public ExcelDataEnum type() default ExcelDataEnum.WORD;

    /**
     * 所属表头名字，没有默认为属性名
     */
    public String name() default  "";

    /**
     * 是否需要和并内容相同的横向单元格
     */
    public boolean mergeVertical() default false;

}
