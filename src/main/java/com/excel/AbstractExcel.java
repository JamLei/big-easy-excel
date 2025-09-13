package com.excel;

import com.excel.anno.ExcelAnno;
import com.excel.enums.Type;
import com.excel.exception.ExcelException;
import com.excel.exception.FileNotFindException;
import com.excel.exception.StreamCreateException;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.io.*;
import java.lang.reflect.Field;
import java.nio.file.Files;
import java.util.List;

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
    protected OutputStream getOutputStream(String fileName) {
        File file = new File(fileName);
        ;
        try {
            return new BufferedOutputStream(Files.newOutputStream(file.toPath()));
        } catch (Exception e) {
            throw new StreamCreateException("file steam crate fail");
        }
    }

    /**
     * 获取文件输入流
     * Get the file input stream
     *
     * @param fileName 文件路径(File path)
     * @return InputStream
     */
    protected InputStream getInputStream(String fileName) {
        File file = new File(fileName);
        if (!file.exists()) {
            throw new FileNotFindException("file not found");
        }
        try {
            return new BufferedInputStream(Files.newInputStream(file.toPath()));
        } catch (Exception e) {
            throw new StreamCreateException("file steam crate fail");
        }
    }


    protected <T> void write(Workbook workbook, List<T> dataList) {
        try {
            Sheet sheet = workbook.createSheet(this.sheet);
            Field[] declaredFields = dataClass.getDeclaredFields();
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < declaredFields.length; i++) {
                ExcelAnno annotation = declaredFields[i].getAnnotation(ExcelAnno.class);
                Cell cell = headerRow.createCell(i);
                CellStyle cellStyle = cell.getCellStyle();
                if (annotation != null) {
                    if (annotation.height() >= 0) {
                        headerRow.setHeight(annotation.height());
                    }
                    cellStyle.setFillBackgroundColor(IndexedColors.GREEN.getIndex());
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
                    Object object = declaredFields[j].get(data);
                    String content = object == null ? "" : Type.getType(object.getClass()).format(object, annotation.format());
                    if (annotation != null) {
                        if (annotation.column() >= 0) {
                            row.createCell(annotation.column()).setCellValue(content);
                        } else {
                            row.createCell(j).setCellValue(content);
                        }
                    }
                }
            }
        } catch (IllegalAccessException e) {
            throw new ExcelException(e);
        }
    }


}
