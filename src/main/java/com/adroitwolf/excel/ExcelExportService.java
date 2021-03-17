package com.adroitwolf.excel;

import com.adroitwolf.anno.ExcelEntity;
import com.adroitwolf.entity.*;
import com.adroitwolf.excel.base.AbstractCommService;
import com.adroitwolf.util.ReflectUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Component;
import org.springframework.util.StringUtils;


import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * <pre>ExcelExportUtil</pre>
 *
 * @author adroitwolf 2021年03月16日 09:58
 */
@Component
@Slf4j
public class ExcelExportService extends AbstractCommService {

    private List<ExportEntity> entityList = new ArrayList<>();

    private RowStyle rowStyle = new RowStyle();


    // 创建表格
    public void createSheetTable(Workbook workbook, Class<?> sourceClass, Collection<?> entity, ExcelInfo info) {
        // 创建 sheet

        Sheet sheet;
        try {
            sheet = workbook.createSheet(info.getSheetName());
        } catch (Exception e) {
            // 如果有重名
            sheet = workbook.createSheet();
        }

        // 获取需要包装的类型
        builderExportEntity(sourceClass);

        // 创建sheet表格
        createSheetContent(sheet, entity, info);

    }

    private Sheet createSheetContent(Sheet sheet, Collection<?> entity, ExcelInfo info) {
        int rowId = 0; // 默认开始

        rowStyle = info.getRowStyle();
        // 首先创建表头
        RowParams rowParams = new RowParams();
        rowParams.setRowStyle(rowStyle);
        rowParams.setRowId(rowId);

        rowId = info.isHideTableHeader() ? rowId : createTableHeader(sheet, rowParams);

        info.setSize(rowId + entity.size());
        // 创建具体内容
        createTableContent(sheet, rowId, entity,info);

        return sheet;

    }

    private void createTableContent(Sheet sheet, int rowNum, Collection<?> entity, ExcelInfo info) {

        for (Object t:entity) {
            setRowValue(sheet,t,rowNum,info);
            rowNum++;
        }
    }


    private int setRowValue(Sheet sheet, Object t, int rowNum, ExcelInfo info) {
        Row row = sheet.getRow(rowNum) == null ? sheet.createRow(rowNum) : sheet.getRow(rowNum);
        row = this.setRowStyle(row, rowStyle);
        for (ExportEntity exportEntity : entityList) {
            String value = String.valueOf(ReflectUtil.invokeMethod(t, exportEntity.getGetMethod()));
            int beforeRowNum = exportEntity.getStartNum() <0 ? rowNum -1 : exportEntity.getStartNum();
            // 如果相同类型合并
            //先要判断一下 和上面是否相等
            if (exportEntity.isMergeVertical() && sheet.getRow(beforeRowNum).getCell(exportEntity.getIndex()) != null) { // 上一行是空的？
                if(sheet.getRow(beforeRowNum).getCell(exportEntity.getIndex()).getStringCellValue().equals(value) && rowNum != info.getSize()-1){ //不相等

                    exportEntity.setStartNum(beforeRowNum);
                }else if(rowNum == info.getSize()-1){
                    CellRangeAddress rangeAddress = new CellRangeAddress(beforeRowNum, rowNum,
                            exportEntity.getIndex(),
                            exportEntity.getIndex());
                    sheet.validateMergedRegions();
                    sheet.addMergedRegion(rangeAddress);
                } else if(beforeRowNum != rowNum-1 ){
                    CellRangeAddress rangeAddress = new CellRangeAddress(beforeRowNum, rowNum-1,
                            exportEntity.getIndex(),
                            exportEntity.getIndex());
                    sheet.validateMergedRegions();
                    sheet.addMergedRegion(rangeAddress);
                    exportEntity.setStartNum(-1);
                    row.createCell(exportEntity.getIndex()).setCellValue(value);
                }else{
                    row.createCell(exportEntity.getIndex()).setCellValue(value);
                }
            } else {
                row.createCell(exportEntity.getIndex()).setCellValue(value);
            }

        }
        return 0;
    }


    private int createTableHeader(Sheet sheet, RowParams rowInfo) {
        Row row = sheet.getRow(rowInfo.getRowId()) == null ? sheet.createRow(rowInfo.getRowId()) : sheet.getRow(rowInfo.getRowId());
        row = this.setRowStyle(row, rowInfo.getRowStyle());
        for (ExportEntity item : entityList) {
            // 设置cellstyle
            //设置内容
            row.createCell(item.getIndex()).setCellValue(item.getName());
        }


        return rowInfo.getRowId() + 1;
    }

    /**
     * 包装ExportEntity
     */
    @Override
    public List<ExportEntity> builderExportEntity(Class<?> sourceClass) { // 这里可以设置row的类型
        // 获取字段
        Field[] fields = sourceClass.getDeclaredFields();
        this.entityList = new ArrayList<>();

        for (int index = 0; index < fields.length; index++) {
            ExcelEntity excelEntity = fields[index].getAnnotation(ExcelEntity.class);
            if (excelEntity == null) { // 如果没有该注解，则设为过滤属性
                continue;
            }
            ExportEntity entity = new ExportEntity();
            entity.setField(fields[index]);
            entity.setName(StringUtils.isEmpty(excelEntity.name()) ? fields[index].getName() : excelEntity.name());
            entity.setIndex(index);
            entity.setMergeVertical(excelEntity.mergeVertical());
            // 设置col样式
            CellStyleEntity cellStyleEntity = new CellStyleEntity();
            cellStyleEntity.setWidth((int) (excelEntity.width() * 2048 / 8.43F));
            entity.setCellStyleEntity(cellStyleEntity);

            // 获取get方法
            entity.setGetMethod(ReflectUtil.getMethod(sourceClass, fields[index].getName()));

            entityList.add(entity);
        }
        return entityList;
    }

}
