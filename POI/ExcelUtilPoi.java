package com.cl.learn.demo;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.util.StrUtil;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;

public class ExcelUtilPoi {

    public void export(Class c, List data) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException, IOException {

        //字段别名
        LinkedHashMap<String, String> columnName = new LinkedHashMap<>();

        //获取需要导出的字段
        Field[] field = c.getDeclaredFields();
        for (int i = 0; i < field.length; i++) {
            ExportCloumn asName = field[i].getAnnotation(ExportCloumn.class);
            if (null != asName) {
                if (StrUtil.isEmpty(asName.asName())) {
                    columnName.put(field[i].getName(), field[i].getName());
                } else {
                    columnName.put(field[i].getName(), asName.asName());
                }
            }
        }
        //创建excel
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("SheetName");

        for (int j = 0; j <= data.size(); j++) {
            HSSFRow row = sheet.createRow(j);
            int i = 0;
            //创建表头
            if (j == 0) {
                for (String key : columnName.keySet()) {
                    sheet.autoSizeColumn(i);
                    sheet.setColumnWidth(i, sheet.getColumnWidth(i) * 35 / 10);
                    row.setHeight((short) 400);
                    HSSFCell cell = row.createCell(i);
                    cell.setCellValue(columnName.get(key));
                    cell.setCellStyle(headStyle(workbook));
                    i++;
                }
            } else {
                for (String key : columnName.keySet()) {
                    row.setHeight((short) 350);
                    HSSFCell cell = row.createCell(i);
                    Object object = data.get(j - 1);
                    String columnC = toUpperCaseFirstOne(key);
                    Method method = object.getClass().getDeclaredMethod("get" + columnC, null);
                    Object value = method.invoke(object);
                    if (null != value) {
                       /* if(key.equals("createdTime") || key.equals("updatedTime")){
                            cell.setCellValue(DateUtil.);
                        }*/
                        if (value instanceof String) {
                            cell.setCellValue(value.toString());
                        } else if (value instanceof Date) {
                            Date date = (Date) value;
                            cell.setCellValue(DateUtil.format(date, "yyyy-MM-dd"));
                        } else if (value instanceof Integer) {
                            cell.setCellValue((Integer) value);
                        } else if (value instanceof Double) {
                            cell.setCellValue((Double) value);
                        } else if (value instanceof Boolean) {
                            cell.setCellValue((Boolean) value);
                        } else if (value instanceof Float) {
                            cell.setCellValue((Float) value);
                        } else if (value instanceof Short) {
                            cell.setCellValue((Short) value);
                        } else if (value instanceof Character) {
                            cell.setCellValue((Character) value);
                        }
                    }
                    i++;
                }
                data.get(j - 1);
            }
        }
        sheet.autoSizeColumn(1);
        OutputStream outputStream = new FileOutputStream("D://test3.xls");
        workbook.write(outputStream);
        outputStream.close();

    }

    public static String toUpperCaseFirstOne(String s) {
        if (Character.isUpperCase(s.charAt(0)))
            return s;
        else
            return (new StringBuilder()).append(Character.toUpperCase(s.charAt(0))).append(s.substring(1)).toString();
    }

    //获取表头样式
    public HSSFCellStyle headStyle(HSSFWorkbook workbook){
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setFillForegroundColor(IndexedColors.SKY_BLUE.index);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        HSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);//字体增粗
        cellStyle.setFont(font);
        return cellStyle;
    }

    //获取表体样式
    public HSSFCellStyle bodyStyle(HSSFWorkbook workbook){
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        HSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 10);
        cellStyle.setFont(font);
        return cellStyle;
    }

    @Test
    public void test() throws NoSuchMethodException, IOException, IllegalAccessException, InvocationTargetException {
        List<Region> regionList = new ArrayList<Region>();
        for (int i = 0; i < 10; i++) {
            Region region = new Region();
            region.setRegionName("地区" + i);
            region.setRegionCode(i + "");
            region.setRegionShortName("地区缩写" + i);
            region.setUpdatedTime(System.currentTimeMillis());
            region.setRegionLevel(i);
            regionList.add(region);
        }
        export(Region.class, regionList);
    }
}
