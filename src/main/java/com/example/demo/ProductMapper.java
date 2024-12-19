package com.example.demo;

import java.util.List;


public interface ProductMapper {
    public List<Product> selectList(List<String> ids);
}
