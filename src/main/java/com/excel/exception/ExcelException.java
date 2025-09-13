package com.excel.exception;

public class ExcelException extends RuntimeException{

    public ExcelException(String message) {
        super(message);
    }

    public ExcelException(Exception e) {
        super(e);
    }
}
