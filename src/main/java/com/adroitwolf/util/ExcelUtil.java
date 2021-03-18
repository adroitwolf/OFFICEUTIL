package com.adroitwolf.util;

import com.adroitwolf.entity.ExcelInfo;
import com.adroitwolf.excel.ExcelExportService;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Collection;

/**
 * <pre>ExcelExportUtil</pre>
 *
 * @author adroitwolf 2021年03月17日 08:08
 */
@Slf4j
public class ExcelUtil {

    private ExcelUtil() {
    }

    private static final ExcelExportService exportService = new ExcelExportService();


    /**
     * 5w以内的数据可以弄，目前不能保存文件
     * @param info 这个可以为null
     */
    public static Workbook exportExcel(Collection<?> list, Class<?> sourceClass, ExcelInfo info,Workbook workbook){
        info = info == null ? exportService.builderInfoByAnno(sourceClass) : info ;
        info.setSize(list.size());
        workbook = workbook == null? exportService.getWorkBook(info): workbook;
        // 获取表头注解
        if(workbook == null){  //有错误发生
            log.error("没有合适的工作表");
            return null;
        }
        //填写内容
        exportService.createSheetTable(workbook, sourceClass, list, info);

        return workbook;
    }

    public static void export2File(String path,Workbook workbook){
        // 创建文件
        exportService.export2File(path,workbook);

    }


}
