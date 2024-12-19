package com.example.demo;


import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.TestInstance;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.Mockito;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.math.BigDecimal;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertNotNull;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@TestInstance(TestInstance.Lifecycle.PER_CLASS)
@Slf4j
public class ProductServiceMockTest {

    @InjectMocks
    ProductService productService;

    @Mock
    ProductMapper productMapper;

    public void getProductsTest(){
        Product product1=new Product();
        product1.setName("");
        product1.setFactory("");
        product1.setPrice(new BigDecimal("12.20"));

        //Mock测试数据
        Mockito.doReturn(List.of(product1)).when(productMapper).selectList(List.of("1","2"));

        //断言
        List<Book> products = productService.selectList();
        assertNotNull(products);
    }

}
