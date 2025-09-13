package com.excel;

import java.util.List;

/**
 * excel interface
 *
 * @author heng.lei
 */
public interface ExcelInterface {

    /**
     * 往excel中写数据
     * Write data to Excel
     *
     * @param data 数据(data)
     * @param <T>  object type
     */
    <T> void doWrite(List<T> data);

    /**
     * 从excel中读取数据
     * Read data from Excel
     *
     * @param clazz 数据对应的实体类(The entity class that the data corresponds to)
     * @param <R>   object type
     * @return object
     */
    <R> List<R> doRead(Class<R> clazz);


}
