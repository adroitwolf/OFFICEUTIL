package com.adroitwolf.anno;

import com.adroitwolf.enums.ExcelWorkEnum;
import com.sun.istack.internal.NotNull;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * <pre>Excel</pre>
 *
 * @author adroitwolf 2021年03月16日 09:44
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE )
public @interface Excel {
    /**
     * 是否隐藏表头，默认false
     */
    public boolean isHideTableFiled() default false;

    public int rowHeight() default 1; //表头默认占多高

    public int rowNum() default 1; // 表头起始行号

    /**
     * 文件存储的路径
     */
    @NotNull
    public String path() ;

    @NotNull
    public ExcelWorkEnum type();

    public short height() default 5; // 这里的高不是像素点

    /**
     * 表格名字
     * @return
     */
    @NotNull
    public String sheetName();

    /**
     *  数据占多高
     */
    public int dataRowHeight() default  1;
}
