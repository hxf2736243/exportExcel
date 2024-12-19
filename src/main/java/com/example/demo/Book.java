package com.example.demo;


import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.math.BigDecimal;

@Data
public class Book {
    @ExcelProperty(value = {"图书统计表","出版社"})
    private String publisher;
    @ExcelProperty(value = {"图书统计表","作者"})
    private String author;
    @ExcelProperty(value = {"图书统计表","书名"})
    private String name;
    @ExcelProperty(value = {"图书统计表","数量"})
    private Integer num;
    @ExcelProperty(value = {"图书统计表","价格"})
    private BigDecimal price;
}
