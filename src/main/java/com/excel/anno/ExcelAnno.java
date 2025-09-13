package com.excel.anno;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.lang.annotation.*;

/**
 * excel实体类注解
 * excel object anno
 *
 * @author heng.lei
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelAnno {

    /**
     * 列名称
     * column name
     */
    String value();


    /**
     * -1代表不通过列匹配
     * -1 means no column matching
     */
    int column() default -1;

    /**
     * format
     */
    String format() default "";

    /**
     * 高度
     * height
     */
    short height() default -1;

    /**
     * width
     */
    int width() default -1;

    /**
     * backgroundColor
     */
    IndexedColors backgroundColor() default IndexedColors.WHITE;

    /**
     * 是否水平居中
     * Whether the level is centered
     */
    boolean isAlignment() default false;

    /**
     * 是否垂直居中
     * Whether it is vertically centered
     */
    boolean isVerticalAlignment() default false;


}
