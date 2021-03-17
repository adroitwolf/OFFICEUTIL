package com.adroitwolf.excel;

import com.adroitwolf.anno.ExcelEntity;
import com.adroitwolf.entity.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import org.springframework.util.StringUtils;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
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
public class ExcelExportService {

    @Autowired
    ExcelCommonService excelCommonService = new ExcelCommonService();


    private List<ExportEntity> entityList = new ArrayList<>();

    private RowStyle rowStyle  = new RowStyle();

    // 创建表格
    public void createSheetTable(Workbook workbook, Class<?> sourceClass, Collection<?> entity, ExcelInfo info) {
        // 创建 sheet

        Sheet sheet;
        try{
            sheet =workbook.createSheet(info.getSheetName());
        } catch (Exception e){
            // 如果有重名
            sheet = workbook.createSheet();
        }

        // 获取需要包装的类型
        buildExportEntityList(sourceClass);

        // 创建sheet表格
        createSheetContent(sheet, entity, info);

    }

    private Sheet createSheetContent(Sheet sheet, Collection<?> entity,  ExcelInfo info) {
        int rowId = 0; // 默认开始

        rowStyle = info.getRowStyle();

        // 首先创建表头
        RowParams rowParams = new RowParams();
        rowParams.setRowStyle(rowStyle);
        rowParams.setRowId(rowId);

        rowId = info.isHideTableHeader() ? rowId : createTableHeader(sheet, rowParams);


        // 创建具体内容
        createTableContent(sheet, rowId, entity, entityList);

        return sheet;

    }

    private void createTableContent(Sheet sheet, int rowId, Collection<?> entity, List<ExportEntity> entityList) {

        for (Object item : entity) {
            Row row = sheet.getRow(rowId) == null ? sheet.createRow(rowId) : sheet.getRow(rowId);
            row = excelCommonService.setRowStyle(row,rowStyle);
            for(ExportEntity exportEntity:entityList){
                try {
                    row.createCell(exportEntity.getIndex()).setCellValue(String.valueOf(exportEntity.getGetMethod().invoke(item)));
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
            rowId++;
        }
    }


    private int createTableHeader(Sheet sheet,  RowParams rowInfo) {
        Row row = sheet.getRow(rowInfo.getRowId()) == null ? sheet.createRow(rowInfo.getRowId()) : sheet.getRow(rowInfo.getRowId());
        row = excelCommonService.setRowStyle(row,rowInfo.getRowStyle());
        for(ExportEntity item :entityList) {
            // 设置cellstyle
            //设置内容
            row.createCell(item.getIndex()).setCellValue(item.getName());
        }

        return rowInfo.getRowId() + 1;
    }

    /**
     * 包装ExportEntity
     */
    private void buildExportEntityList(Class<?> sourceClass) { // 这里可以设置row的类型
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


            // 设置col样式
            CellStyleEntity cellStyleEntity = new CellStyleEntity();
            cellStyleEntity.setWidth((int) (excelEntity.width() * 2048 / 8.43F));
            entity.setCellStyleEntity(cellStyleEntity);

            PropertyDescriptor descriptor = null;
            // 获取get方法
            try {
                descriptor = new PropertyDescriptor(fields[index].getName(), sourceClass);
                entity.setGetMethod(descriptor.getReadMethod());
            } catch (IntrospectionException e) {
                e.printStackTrace();
            }
            entityList.add(entity);
        }
    }

}
