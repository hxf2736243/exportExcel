package com.example.demo;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.math.BigDecimal;


@Data
public class Product {
    @ExcelProperty("产品名称")
    private String name;
    @ExcelProperty("生产厂家")
    private String factory;
    @ExcelProperty("产品价格")
    private BigDecimal price;
}
