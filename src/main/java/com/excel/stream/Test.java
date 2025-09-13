package com.excel.stream;

import com.excel.Excel;
import com.excel.anno.ExcelAnno;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.ArrayList;
import java.util.List;

public class Test {

    public static void main(String[] args) {
        List<Test.Person> list = new ArrayList<>();
        list.add(new Person("张三", 22, "江西"));
        list.add(new Person("张三", 22, "江西"));
        list.add(new Person("张三", 22, "江西"));
        list.add(new Person("张三", 22, "江西"));
        list.add(new Person("张三", 22, "江西"));
        list.add(new Person("张三", 22, "江西"));
        list.add(new Person("张三", 22, "江西"));
        Excel.dataType(Test.Person.class).fileName("D:\\upload\\google\\测试1.xlsx").sheet("测试").doWrite(list);
    }


    public static class Person {

        @ExcelAnno(value = "姓名",backgroundColor = IndexedColors.BLUE1)
        private String name;

        @ExcelAnno("年龄")
        private Integer age;

        @ExcelAnno("地址")
        private String address;

        public Person(String name, int age, String address) {
            this.name = name;
            this.age = age;
            this.address = address;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public int getAge() {
            return age;
        }

        public void setAge(int age) {
            this.age = age;
        }

        public String getAddress() {
            return address;
        }

        public void setAddress(String address) {
            this.address = address;
        }
    }
}
