package com.adroitwolf.excel.base;

import com.adroitwolf.anno.Excel;
import com.adroitwolf.entity.ExcelInfo;
import com.adroitwolf.entity.ExportEntity;
import com.adroitwolf.entity.RowStyle;
import com.adroitwolf.enums.ExcelWorkEnum;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

/**
 * <pre>ExcelCommonService</pre>
 *
 * @author adroitwolf 2021年03月17日 08:12
 */
@Slf4j
@Component
public abstract class AbstractCommService {

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



        info.setHideTableHeader(excel.isHideTableFiled());

        info.setStartRow(excel.rowNum() ); // excel从0开始，但是用户喜欢从1开始，后面需要规定好

        info.setHeaderHeight(excel.rowHeight());

        info.setDataRowHeight(excel.dataRowHeight());

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

    public void setRowStyle(Row row, RowStyle rowStyle){
        if(rowStyle != null){
            row.setHeight(rowStyle.getHeight());
        }
    }

    public abstract  List<ExportEntity> builderExportEntity(Class<?> sourceClass);



    /**
     *  设定单元值
     */
    public  void setCellValue(int startRow, int endRow, int startCol, int endCol, Sheet sheet, Object t, RowStyle rowStyle){
        Row row = sheet.getRow(startRow) == null ? sheet.createRow(startRow) : sheet.getRow(startRow);
        for(int i = startRow;i<=endRow;i++){
            Row styleRow = sheet.getRow(i) == null ? sheet.createRow(i) : sheet.getRow(i);
            setRowStyle(styleRow,rowStyle);
        }

        if(startRow != endRow || startCol != endCol){ //说明这是一个需要合并的单元

            CellRangeAddress rangeAddress = new CellRangeAddress(startRow, endRow,
                    startCol,
                    endCol);

            sheet.validateMergedRegions();

            sheet.addMergedRegion(rangeAddress);
        }

        Cell cell = row.getCell(startCol) == null ? row.createCell(startCol): row.getCell(startCol);

        cell.setCellValue(String.valueOf(t));
    }

}
