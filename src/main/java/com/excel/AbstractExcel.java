package com.excel;

import com.excel.anno.ExcelAnno;
import com.excel.enums.Type;
import com.excel.exception.ExcelException;
import com.excel.exception.FileNotFindException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 抽象excel接口实现类
 * Abstract Excel interface implementation class
 *
 * @author heng.lei
 */
public abstract class AbstractExcel implements ExcelInterface {

    /**
     * xls
     */
    protected final String xls = ".xls";

    /**
     * xlsx
     */
    protected final String xlsx = ".xlsx";


    /**
     * sheet
     */
    protected String sheet;

    /**
     * 文件全路径
     * File full path
     */
    protected String fileName;

    /**
     * 数据类型
     * <p>
     * dataType
     */
    protected Class<?> dataClass;


    /**
     * 获取文件输出流
     * Get the file output stream
     *
     * @param fileName 文件路径(File path)
     * @return OutputStream
     */
    protected abstract OutputStream getOutputStream(String fileName);

    /**
     * 获取文件输入流
     * Get the file input stream
     *
     * @param fileName 文件路径(File path)
     * @return InputStream
     */
    protected abstract InputStream getInputStream(String fileName);

    /**
     * 将实体类中的数据写到excel中
     * Write the data in the entity class into Excel
     *
     * @param dataList 实体类数据(object)
     * @param <T>      object Type
     */
    protected <T> void write(List<T> dataList) {
        try (Workbook workbook = new XSSFWorkbook(); OutputStream outStream = getOutputStream(fileName)) {
            Sheet sheet = workbook.createSheet(this.sheet);
            Field[] declaredFields = dataClass.getDeclaredFields();
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < declaredFields.length; i++) {
                ExcelAnno annotation = declaredFields[i].getAnnotation(ExcelAnno.class);
                Cell cell = headerRow.createCell(i);
                CellStyle cellStyle = cell.getSheet().getWorkbook().createCellStyle();
                if (annotation != null) {
                    cellStyle.setFillForegroundColor(annotation.backgroundColor().index);
                    cellStyle.setFillPattern(annotation.fillPatternType());
                    cellStyle.setBorderTop(BorderStyle.THIN);
                    cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
                    cellStyle.setBorderRight(BorderStyle.THIN);
                    cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
                    cellStyle.setBorderBottom(BorderStyle.THIN);
                    cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
                    cellStyle.setBorderLeft(BorderStyle.THIN);
                    cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
                    if (annotation.height() >= 0) {
                        headerRow.setHeight(annotation.height());
                    }
                    if (annotation.width() >= 0) {
                        sheet.setColumnWidth(i, annotation.width());
                    }
                    if (annotation.isAlignment()) {
                        cellStyle.setAlignment(HorizontalAlignment.CENTER);
                    }
                    if (annotation.isVerticalAlignment()) {
                        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                    }
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue(annotation.value());
                }
            }
            for (int i = 0; i < dataList.size(); i++) {
                T data = dataList.get(i);
                Row row = sheet.createRow(i + 1);
                for (int j = 0; j < declaredFields.length; j++) {
                    declaredFields[j].setAccessible(true);
                    ExcelAnno annotation = declaredFields[j].getAnnotation(ExcelAnno.class);
                    if (annotation != null) {
                        Object object = declaredFields[j].get(data);
                        String content = object == null ? "" : Type.getType(declaredFields[j].getType()).format(object, annotation.format());
                        if (annotation.column() >= 0) {
                            row.createCell(annotation.column()).setCellValue(content);
                        } else {
                            row.createCell(j).setCellValue(content);
                        }
                    }

                }
            }
            workbook.write(outStream);
        } catch (Exception e) {
            throw new ExcelException(e);
        }
    }

    /**
     * 将excel数据装载到实体类中
     * Load Excel data into entity classes
     *
     * @param clazz 实体类class
     * @param <R>   object type
     * @return 实体类集合(object list)
     */
    protected <R> List<R> read(Class<R> clazz) {
        List<R> list = new ArrayList<>();
        try (InputStream inputStream = getInputStream(fileName);
             Workbook workbook = WorkbookFactory.create(inputStream)) {
            Sheet sheet = workbook.getSheet(this.sheet);
            if (sheet == null) {
                throw new ExcelException("Sheet not found");
            }
            Row header = sheet.getRow(0);
            if (header == null) {
                return list;
            }
            int physicalNumberOfCells = header.getPhysicalNumberOfCells();
            Map<String, Integer> map = new HashMap<>();
            for (int i = 0; i < physicalNumberOfCells; i++) {
                Cell cell = header.getCell(i);
                if (cell == null) {
                    continue;
                }
                map.put(cell.toString(), i);
            }
            int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
            Field[] declaredFields = clazz.getDeclaredFields();
            for (int i = 1; i < physicalNumberOfRows; i++) {
                Row row = sheet.getRow(i);
                R object = clazz.newInstance();
                for (Field declaredField : declaredFields) {
                    declaredField.setAccessible(true);
                    Class<?> type = declaredField.getType();
                    Type contentType = Type.getType(type);
                    Object value = null;
                    ExcelAnno annotation = declaredField.getAnnotation(ExcelAnno.class);
                    if (annotation == null) {
                        continue;
                    }
                    int column = annotation.column();
                    if (column >= 0) {
                        Cell cell = row.getCell(column);
                        value = contentType.parse(cell.toString(), annotation.format());
                    } else {
                        Integer index = map.get(annotation.value());
                        if (index != null) {
                            Cell cell = row.getCell(index);
                            value = contentType.parse(cell.toString(), annotation.format());
                        }
                    }
                    declaredField.set(object, value);
                }
                list.add(object);
            }
        } catch (Exception e) {
            throw new ExcelException(e);
        }
        return list;
    }


    protected File checkFileNameAndCreateFile(String fileName) {
        if (fileName.endsWith(xls) || fileName.endsWith(xlsx)) {
            return new File(fileName);
        }
        throw new FileNotFindException("file path exception");
    }


}
