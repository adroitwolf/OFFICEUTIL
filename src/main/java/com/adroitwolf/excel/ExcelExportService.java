package com.adroitwolf.excel;

import com.adroitwolf.anno.ExcelEntity;
import com.adroitwolf.entity.*;
import com.adroitwolf.excel.base.AbstractCommService;
import com.adroitwolf.util.ReflectUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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
        entityList = builderExportEntity(sourceClass);

        // 创建sheet表格
        createSheetContent(sheet, entity, info);

    }

    private Sheet createSheetContent(Sheet sheet, Collection<?> entity, ExcelInfo info) {
        int rowId = info.getStartRow() - 1; // 默认开始

        rowId = info.isHideTableHeader() ? rowId : createTableHeader(sheet, info);

        // 创建具体内容
        createTableContent(sheet, rowId, entity, info);

        return sheet;

    }

    private void createTableContent(Sheet sheet, int rowNum, Collection<?> entity, ExcelInfo info) {
        int currentSize = 0; // 记录当前数据的索引
        for (Object t : entity) {
            currentSize ++;
            for (ExportEntity exportEntity : entityList) {


                Object value = ReflectUtil.invokeMethod(t, exportEntity.getGetMethod());

                if (!exportEntity.isMergeVertical()) { // 如果是用户规定了长度
                    setCellValue(rowNum, rowNum + info.getDataRowHeight() - 1, exportEntity.getIndex(), exportEntity.getEndNum(), sheet, value, info.getRowStyle());
                } else {
                    if(exportEntity.getBeforeValue() == null){ // 说明是第一列
                        // 直接添加
                        exportEntity.setBeforeValue(value);
                        exportEntity.setStartRowNum(rowNum);
                    }else if(currentSize == info.getSize()){ // 如果和上面那个相等并且还是最后一个
                        if(!exportEntity.getBeforeValue().equals(value)){ //添加上一个
                            setCellValue(exportEntity.getStartRowNum(), rowNum + info.getDataRowHeight() - 2, exportEntity.getIndex(), exportEntity.getEndNum(), sheet, exportEntity.getBeforeValue(), info.getRowStyle());
                            setCellValue(rowNum,rowNum + info.getDataRowHeight() - 1,exportEntity.getIndex(), exportEntity.getEndNum(), sheet, exportEntity.getBeforeValue(), info.getRowStyle());
                        }else{
                            setCellValue(exportEntity.getStartRowNum(), rowNum + info.getDataRowHeight() - 1, exportEntity.getIndex(), exportEntity.getEndNum(), sheet, value, info.getRowStyle());
                        }
                    }else if(!exportEntity.getBeforeValue().equals(value)){
                        setCellValue(exportEntity.getStartRowNum(), rowNum + info.getDataRowHeight() - 2, exportEntity.getIndex(), exportEntity.getEndNum(), sheet, exportEntity.getBeforeValue(), info.getRowStyle());
                        exportEntity.setBeforeValue(value);
                        exportEntity.setStartRowNum(rowNum);
                    }
                }
            }
            rowNum += info.getDataRowHeight() ;

        }
    }




    private int createTableHeader(Sheet sheet, ExcelInfo rowInfo) {
        for (ExportEntity item : entityList) {
            setCellValue(rowInfo.getStartRow() - 1, rowInfo.getHeaderHeight(), item.getIndex(), item.getEndNum(), sheet, item.getName(), rowInfo.getRowStyle());
        }
        return rowInfo.getStartRow() + rowInfo.getHeaderHeight() - 1;
    }

    /**
     * 包装ExportEntity
     */
    @Override
    public List<ExportEntity> builderExportEntity(Class<?> sourceClass) { // 这里可以设置row的类型
        // 获取字段
        Field[] fields = sourceClass.getDeclaredFields();

        List<ExportEntity> entityList = new ArrayList<>();

        int index = 0; // 标题的索引列坐标

        for (Field field : fields) {

            ExcelEntity excelEntity = field.getAnnotation(ExcelEntity.class);
            if (excelEntity == null) { // 如果没有该注解，则设为过滤属性
                continue;
            }

            ExportEntity entity = new ExportEntity();
            entity.setName(StringUtils.isEmpty(excelEntity.name()) ? field.getName() : excelEntity.name());
            entity.setMergeVertical(excelEntity.mergeVertical());
            // 设置col样式
            CellStyleEntity cellStyleEntity = new CellStyleEntity();
            cellStyleEntity.setWidth((int) (excelEntity.width() * 2048 / 8.43F));
            entity.setCellStyleEntity(cellStyleEntity);

            // 获取get方法
            entity.setGetMethod(ReflectUtil.getMethod(sourceClass, field.getName()));

            entity.setIndex(index);

            // 最终的列索引
            entity.setEndNum(index + excelEntity.cellWidth() - 1);


            entityList.add(entity);

            index += excelEntity.cellWidth();
        }
        return entityList;
    }

}
