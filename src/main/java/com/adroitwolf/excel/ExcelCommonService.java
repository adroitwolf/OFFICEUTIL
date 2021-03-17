package com.adroitwolf.excel;

import com.adroitwolf.anno.Excel;
import com.adroitwolf.entity.CellStyleEntity;
import com.adroitwolf.entity.ExcelInfo;
import com.adroitwolf.entity.RowStyle;
import com.adroitwolf.enums.ExcelWorkEnum;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileOutputStream;

/**
 * <pre>ExcelCommonService</pre>
 *
 * @author adroitwolf 2021年03月17日 08:12
 */
@Slf4j
@Component
public class ExcelCommonService {

    public Workbook getWorkBook( ExcelInfo info){
        ExcelWorkEnum type;
        if(null != info){
            type = info.getType();
            return getWorkbookByType(type);
        }else{
            log.error("必要表格数据丢失，请核查后重试");
            return null;
        }
    }
    private Workbook getWorkbookByType(ExcelWorkEnum type){
        switch (type){
            case HSSF:
                return new HSSFWorkbook();
            case XSSF:
                return new XSSFWorkbook();
            case SXSSF:
            default:
                return new SXSSFWorkbook();
        }

    }
    public  ExcelInfo builderInfoByAnno(Class<?> sourceClass){
        if(!sourceClass.isAnnotationPresent(Excel.class)){
            log.error("必要表格数据丢失，请核查后重试");
            return null;
        }
        Excel excel = sourceClass.getAnnotation(Excel.class);

        ExcelInfo info = new ExcelInfo();
        info.setSheetName(excel.sheetName());

        info.setType(excel.type());

//        info.setTableName(new StringBuilder()
//                .append(excel.name())
//                .append(excel.type().getSuffix())
//                .toString());
        info.setHideTableHeader(excel.isHideTableFiled());

        info.setPath(excel.path());

        RowStyle style = new RowStyle();
        style.setHeight((short) (excel.height()  * 2048 / 8.43F));

        info.setRowStyle(style);

        return info;
    }

    public void export2File(String path, Workbook workbook){
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(new File(path));
            workbook.write(out);
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public Row setRowStyle(Row row, RowStyle rowStyle){
        row.setHeight(rowStyle.getHeight());
        return row;
    }

//    public Cell setCellStyle(Cell cell, CellStyleEntity cellStyleEntity){
//
//    }
}
