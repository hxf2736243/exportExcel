package com.boxer.commom.annotation;


import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExportExcel {
    String fileName() default "data.xlsx";
    int[] sameValueMergeColumns() default {};

    int headHeight() default 1;
    int totalMergeColumnStart() default 0;
    int totalMergeColumnEnd() default 0;
    String totalText() default "合计";
}
