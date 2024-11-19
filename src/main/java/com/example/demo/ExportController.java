package com.example.demo;


import com.boxer.commom.annotation.ExportExcel;
import com.boxer.commom.annotation.LogExecutionTime;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
@RestController
public class ExportController {

    @ExportExcel()
    @GetMapping("/book/export")
    public List<Book> export(){
        List<Book> bookList=new ArrayList<>();
        Book book=new Book();
        book.setName("钢铁是怎样炼成的");
        book.setAuthor("保尔柯察金");
        bookList.add(book);
        return bookList;
    }

    @ExportExcel()
    @GetMapping("/product/export")
    public List<Product> exportProduct(){
        List<Product> bookList=new ArrayList<>();
        Product book=new Product();
        book.setName("钢铁是怎样炼成的");
        book.setFactory("机械出版社");
        book.setPrice(new BigDecimal("12.5"));
        bookList.add(book);
        return bookList;
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
