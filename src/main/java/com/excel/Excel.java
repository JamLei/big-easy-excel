package com.excel;

import com.excel.exception.ExcelException;
import com.excel.exception.FileNotFindException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;


/**
 * excel 入口使用类
 * Excel entry using classes
 *
 * @author heng.lei
 */
public final class Excel<D> extends AbstractExcel {

    /**
     * 文件全路径
     * File full path
     */
    private String fileName;


    private Excel(Class<D> dataClass) {
        super.dataClass = dataClass;
    }


    /**
     * 往excel中写数据
     * Write data to Excel
     *
     * @param dataList 数据(data)
     * @param <T>      object type
     */
    @Override
    public <T> void doWrite(List<T> dataList) {
        Workbook workbook = null;
        OutputStream outStream = null;
        try {
            workbook = new XSSFWorkbook();
            write(workbook, dataList);
            outStream = getOutputStream(fileName);
            workbook.write(outStream);
        } catch (Exception e) {
            throw new ExcelException(e);
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {

                }
            }
            if (outStream != null) {
                try {
                    outStream.close();
                } catch (IOException e) {
                }
            }

        }
    }


    /**
     * 从excel中读取数据
     * Read data from Excel
     *
     * @param clazz 数据对应的实体类(The entity class that the data corresponds to)
     * @param <R>   object type
     * @return object
     */
    @Override
    public <R> List<R> doRead(Class<R> clazz) {
        return null;
    }

    /**
     * 指定数据类型
     *
     * @param dataType 数据类型
     * @param <T>      数据类型
     * @return Excel
     */
    public static <T> Excel<T> dataType(Class<T> dataType) {
        return new Excel<>(dataType);
    }

    /**
     * 文件全路径
     * File full path
     *
     * @param fileName 文件全路径(File full path)
     * @return Excel
     */
    public Excel<D> fileName(String fileName) {
        if (fileName != null && !fileName.isEmpty()) {
            if (fileName.endsWith(".xls") || fileName.endsWith(".xlsx")) {
                this.fileName = fileName;
                return this;
            }
        }
        throw new FileNotFindException("file name is error");
    }

    /**
     * sheet name
     *
     * @param sheet sheet
     * @return Excel
     */
    public Excel<D> sheet(String sheet) {
        super.sheet = sheet;
        return this;
    }

}
