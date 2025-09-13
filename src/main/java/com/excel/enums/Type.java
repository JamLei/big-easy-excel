package com.excel.enums;

import com.excel.exception.ExcelException;
import com.excel.utils.DateFormat;

import java.math.BigDecimal;
import java.util.Arrays;
import java.util.Date;

public enum Type {

    DATE(Date.class) {
        @Override
        public String format(Object filed, String format) {
            Date date = (Date) filed;
            return DateFormat.format(date, format);
        }

        @Override
        public Object parse(String value, String format) {
            return DateFormat.parse(value, format);
        }
    },
    BIG_DECIMAL(BigDecimal.class) {
        @Override
        public String format(Object filed, String format) {
            BigDecimal bigDecimal = (BigDecimal) filed;
            return bigDecimal.toString();
        }

        @Override
        public Object parse(String value, String format) {
            return new BigDecimal(value);
        }
    },
    INTEGER(Integer.class) {
        @Override
        public String format(Object filed, String format) {
            return filed.toString();
        }

        @Override
        public Object parse(String value, String format) {
            return Integer.parseInt(value);
        }
    },
    LONG(Long.class) {
        @Override
        public String format(Object filed, String format) {
            return filed.toString();
        }

        @Override
        public Object parse(String value, String format) {
            return Long.parseLong(value);
        }
    },
    BYTE(Byte.class) {
        @Override
        public String format(Object filed, String format) {
            return filed.toString();
        }

        @Override
        public Object parse(String value, String format) {
            return Byte.parseByte(value);
        }
    },
    SHORT(Short.class) {
        @Override
        public String format(Object filed, String format) {
            return filed.toString();
        }

        @Override
        public Object parse(String value, String format) {
            return Short.parseShort(value);
        }
    },
    BOOLEAN(Boolean.class) {
        @Override
        public String format(Object filed, String format) {
            return filed.toString();
        }

        @Override
        public Object parse(String value, String format) {
            return Boolean.parseBoolean(value);
        }
    },
    DOUBLE(Double.class) {
        @Override
        public String format(Object filed, String format) {
            return filed.toString();
        }

        @Override
        public Object parse(String value, String format) {
            return Double.parseDouble(value);
        }
    },
    FLOAT(Float.class) {
        @Override
        public String format(Object filed, String format) {
            return filed.toString();
        }

        @Override
        public Object parse(String value, String format) {
            return Float.parseFloat(value);
        }
    }, CHAR(Character.class) {
        @Override
        public String format(Object filed, String format) {
            return filed.toString();
        }

        @Override
        public Object parse(String value, String format) {
            return value.charAt(0);
        }
    }, STRING(String.class) {
        @Override
        public String format(Object filed, String format) {
            return filed.toString();
        }

        @Override
        public Object parse(String value, String format) {
            return value;
        }
    };

    private final Class<?> fileType;

    Type(Class<?> fileType) {
        this.fileType = fileType;
    }

    public abstract String format(Object file, String format);

    public abstract Object parse(String value, String format);

    public static Type getType(Class<?> filedType) {
        return Arrays.stream(Type.values()).filter(type -> type.fileType == filedType)
                .findFirst()
                .orElseThrow(() -> new ExcelException("class type not find"));
    }

}
