package com.example.demo;

import com.alibaba.excel.annotation.ExcelProperty;
import com.boxer.commom.annotation.MergeExcelColumn;
import lombok.Data;

import java.math.BigDecimal;


@Data
public class Product {
    @ExcelProperty("图书名称")
    private String name;
    @ExcelProperty(value = "出版社",index = 0)
    private String factory;
    @ExcelProperty(value = "作者",index = 0)
    private String author;

    @ExcelProperty("价格")
    private BigDecimal price;
}
