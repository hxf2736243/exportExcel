package com.example.demo;


import com.boxer.commom.annotation.ExportExcel;
import com.boxer.commom.annotation.LogExecutionTime;
import jakarta.annotation.Resource;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;
@RestController
public class ExportController {

    @Resource
    ProductService productService;
    @ExportExcel(fileName = "图书统计表.xlsx",mergeColumns={0,1},totalText = "总计",headHeight = 2,totalRowMergeStart = 0,totalRowMergeEnd = 2)
    @GetMapping("/product/export")
    public List<Book> export(){
        return productService.selectList();
    }


    @LogExecutionTime
    @GetMapping("/detail")
    public Book detail(){
        Book book=new Book();
        book.setName("钢铁是怎样炼成的");
        book.setAuthor("保尔柯察金");
        return book;
    }
}
