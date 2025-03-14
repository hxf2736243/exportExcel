package com.boxer.commom.grok2;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.NumberFormat;
import lombok.Data;

import java.math.BigDecimal;

@Data
public class StudentExportDTO {
    @ExcelProperty(value = {"2025上学期成绩","年级"})
    private String grade;       // 年级（全局合并）
    @ExcelProperty(value = {"2025上学期成绩","学号"})
    private String studentId;   // 学号
    @ExcelProperty(value = {"2025上学期成绩","名称"})
    private String name;        // 姓名（全局合并）
    @ExcelProperty(value = {"2025上学期成绩","课程"})
    private String course;      // 课程（按学号合并）
    @ExcelProperty(value = {"2025上学期成绩","分数"})
    @NumberFormat("0.00")
    private BigDecimal score;       // 分数

    // Getter 和 Setter 省略


    public StudentExportDTO(String studentId, String name, String course, String grade,BigDecimal score) {
        this.studentId = studentId;
        this.name = name;
        this.course = course;
        this.grade = grade;
        this.score=score;
    }
}
