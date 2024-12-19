package com.boxer.commom.annotation;


import java.lang.annotation.*;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface TotalRow {
    boolean enableTotal() default false;
    int headHeight() default 1;
    int totalMergeColumnStart() default 0;
    int totalMergeColumnEnd() default 0;
    String totalText() default "合计";
}
