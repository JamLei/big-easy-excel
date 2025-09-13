package com.excel;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

public final class NetworkExcel<D> extends AbstractExcel {

    private NetworkExcel(Class<D> dataClass) {
        super.dataClass = dataClass;
    }

    private OutputStream out;

    private InputStream in;


    /**
     * 获取文件输出流
     * Get the file output stream
     *
     * @param fileName 文件路径(File path)
     * @return OutputStream
     */
    @Override
    protected OutputStream getOutputStream(String fileName) {
        return out;
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
        return in;
    }

    /**
     * 往excel中写数据
     * Write data to Excel
     *
     * @param data 数据(data)
     * @param <T>  object type
     */
    @Override
    public <T> void doWrite(List<T> data) {
        write(data);
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
    public static <T> NetworkExcel<T> dataType(Class<T> dataType) {
        return new NetworkExcel<>(dataType);
    }


    /**
     * 指定输出流
     * Specify the output stream
     *
     * @param outputStream 输出流(OutputStream)
     * @return NetworkExcel
     */
    public NetworkExcel<D> outputStream(OutputStream outputStream) {
        this.out = outputStream;
        return this;
    }

    /**
     * 指定输入流
     * Specify the input stream
     *
     * @param inputStream 输入流(InputStream)
     * @return NetworkExcel
     */
    public NetworkExcel<D> inputStream(InputStream inputStream) {
        this.in = inputStream;
        return this;
    }

    /**
     * sheet name
     *
     * @param sheet sheet
     * @return Excel
     */
    public NetworkExcel<D> sheet(String sheet) {
        super.sheet = sheet;
        return this;
    }

}
