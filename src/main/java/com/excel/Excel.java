package com.excel;

import com.excel.exception.ExcelException;
import com.excel.exception.FileNotFindException;
import com.excel.exception.StreamCreateException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;


/**
 * 本地excel文件入口使用类
 * Excel entry using classes
 *
 * @author heng.lei
 */
public final class Excel<D> extends AbstractExcel {


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
        write(dataList);
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
        return read(clazz);
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

    /**
     * 获取文件输出流
     * Get the file output stream
     *
     * @param fileName 文件路径(File path)
     * @return OutputStream
     */
    @Override
    protected OutputStream getOutputStream(String fileName) {
        File file = checkFileNameAndCreateFile(fileName);
        try {
            return new BufferedOutputStream(Files.newOutputStream(file.toPath()));
        } catch (Exception e) {
            throw new StreamCreateException("file is reading");
        }
    }


    /**
     * 获取文件输入流
     * Get the file input stream
     *
     * @param fileName 文件路径(File path)
     * @return InputStream
     */
    @Override
    protected InputStream getInputStream(String fileName) {
        File file = checkFileNameAndCreateFile(fileName);
        if (!file.exists()) {
            throw new FileNotFindException("file not found");
        }
        try {
            return new BufferedInputStream(Files.newInputStream(file.toPath()));
        } catch (Exception e) {
            throw new StreamCreateException("file is reading");
        }

    }
}
