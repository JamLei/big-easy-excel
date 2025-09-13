package com.excel.utils;

import com.excel.exception.ExcelException;

import java.text.SimpleDateFormat;
import java.util.Date;

public class DateFormat {

    public static String format(Date date,String format) {
        SimpleDateFormat sdf = new SimpleDateFormat(format);
        return sdf.format(date);
    }


    public static Date parse(String date,String format) {
        try {
            SimpleDateFormat sdf = new SimpleDateFormat(format);
            return sdf.parse(date);
        }catch (Exception e){
            throw new ExcelException(e);
        }
    }
}
