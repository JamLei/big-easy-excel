package com.excel;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

public class NetworkExcel extends AbstractExcel {


    @Override
    protected OutputStream getOutputStream(String fileName) {
        return null;
    }

    @Override
    protected InputStream getInputStream(String fileName) {
        return null;
    }

    @Override
    public <T> void doWrite(List<T> data)  {
            write(data);
    }

    @Override
    public <R> List<R> doRead(Class<R> clazz) {
        return read(clazz);
    }
}
