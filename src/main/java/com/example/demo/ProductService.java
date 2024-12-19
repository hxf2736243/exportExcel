package com.example.demo;


import jakarta.annotation.Resource;
import org.springframework.stereotype.Service;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

@Service
public class ProductService {

    public List<Book> selectList(){
        List<Book> bookList=new ArrayList<>();
        Book book=new Book();
        book.setName("钢铁是怎样炼成的");
        book.setPublisher("机械出版社");
        book.setAuthor("保尔柯察金");
        book.setNum(1);
        book.setPrice(new BigDecimal("12.5"));
        Book book1=new Book();
        book1.setName("java编程思想");
        book1.setPublisher("机械出版社");
        book1.setAuthor("Bruce Eckel");
        book1.setNum(11);
        book1.setPrice(new BigDecimal("122.5"));
        Book book2=new Book();
        book2.setName("红楼梦");
        book2.setPublisher("机械出版社");
        book2.setAuthor("曹雪芹");
        book2.setNum(2);
        book2.setPrice(new BigDecimal("88.5"));

        Book book3=new Book();
        book3.setName("C++编程思想");
        book3.setPublisher("机械出版社");
        book3.setAuthor("Bruce Eckel");
        book3.setNum(8);
        book3.setPrice(new BigDecimal("99.5"));
        bookList.add(book);
        bookList.add(book1);
        bookList.add(book3);
        bookList.add(book2);

        return bookList;
    }

}
