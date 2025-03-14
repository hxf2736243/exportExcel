package com.boxer.commom.grok2;

import com.alibaba.excel.EasyExcel;

import java.math.BigDecimal;
import java.util.Arrays;
import java.util.List;

import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import org.apache.poi.ss.usermodel.*;

public class ExportTest {
    public static void main(String[] args) throws Exception {
        List<StudentExportDTO> dataList = Arrays.asList(
                new StudentExportDTO("S001", "张三", "数学", "高一", new BigDecimal("97.20")),
                new StudentExportDTO("S001", "张三", "英语", "高一", new BigDecimal("99.50")),
                new StudentExportDTO("S002", "李四", "数学", "高一", new BigDecimal("87.00")),
                new StudentExportDTO("S002", "李四", "英语", "高一", new BigDecimal("100.00")),
                new StudentExportDTO("总计", "总计", "总计", "总计", new BigDecimal("398.00"))
        );

        // 定义合并策略
        int[] globalMergeCols = new int[]{0}; // 姓名 (列 1) 和 年级 (列 3) 全局合并
        int[] groupMergeCols = new int[]{1, 2};     // 课程 (列 2) 按学号合并
        int groupByCol = 1;                      // 学号 (列 0) 作为分组依据

        // 创建全局样式
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER); // 水平居中
        contentWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);     // 垂直居中
        contentWriteCellStyle.setBorderTop(BorderStyle.THIN);    // 上边框
        contentWriteCellStyle.setBorderBottom(BorderStyle.THIN); // 下边框
        contentWriteCellStyle.setBorderLeft(BorderStyle.THIN);   // 左边框
        contentWriteCellStyle.setBorderRight(BorderStyle.THIN);  // 右边框

        // 如果需要表头样式不同，可以单独定义
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        headWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        headWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headWriteCellStyle.setBorderTop(BorderStyle.THIN);
        headWriteCellStyle.setBorderBottom(BorderStyle.THIN);
        headWriteCellStyle.setBorderLeft(BorderStyle.THIN);
        headWriteCellStyle.setBorderRight(BorderStyle.THIN);

        // 设置样式策略
        HorizontalCellStyleStrategy styleStrategy = new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);
        // 导出
        String fileName = "students1.xlsx";
        EasyExcel.write(fileName, StudentExportDTO.class)
                .registerWriteHandler(new CustomMergeHandler(dataList,2, globalMergeCols, groupMergeCols, groupByCol,0,3))
                .registerWriteHandler(styleStrategy)
                .sheet("学生信息")
                .doWrite(dataList);
    }


}