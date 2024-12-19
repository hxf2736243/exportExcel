package com.example.demo;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.handler.RowWriteHandler;
import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.handler.WriteHandler;
import com.alibaba.excel.write.handler.context.RowWriteHandlerContext;
import com.alibaba.excel.write.handler.context.SheetWriteHandlerContext;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.WriteTable;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

public class MultiColumnTotalWithMergeExample {

    public static void main(String[] args) {
        String fileName = "total-row-with-merge.xlsx";

        // 准备数据
        List<Data> dataList = generateData();

        // 设置基础样式
        HorizontalCellStyleStrategy horizontalCellStyleStrategy = new HorizontalCellStyleStrategy(getHeadStyle(), getContentStyle());


        // 把sheet设置为不需要头 不然会输出sheet的头 这样看起来第一个table 就有2个头了
        WriteSheet writeSheet = EasyExcel.writerSheet("sheetName").needHead(Boolean.FALSE).build();
        // 这里必须指定需要头，table 会继承sheet的配置，sheet配置了不需要，table 默认也是不需要
        WriteTable writeTable0 = EasyExcel.writerTable(0).needHead(Boolean.TRUE).registerWriteHandler(horizontalCellStyleStrategy).build();


        ExcelWriter excelWriter = EasyExcel.write(fileName, Data.class).autoCloseStream(false)
                .build();

        // 第一次写入会创建头
        excelWriter.write(dataList, writeSheet, writeTable0);

        Data data1 = calculateTotalRow(dataList);

        // 写入合计行并合并单元格,
        // 第二次写不创建头，然后在第一次的后面写入数据
        excelWriter.write(List.of(data1), writeSheet, writeTable0);
        Sheet sheet = excelWriter.writeContext().writeSheetHolder().getSheet();
        sheet.addMergedRegion(new CellRangeAddress(dataList.size()+1, dataList.size()+1, 0, 1)); // 合并单元格
//        cell.setCellValue("合计"); // 设置合计文本

        excelWriter.finish();
    }

    // 数据类
    public static class Data {
        public Data(){}
        private String name;
        private String category;
        private Double value1;
        private Double value2;
        private Double value3;

        public String getCategory() {
            return category;
        }

        public void setCategory(String category) {
            this.category = category;
        }

        public Data(String name,String category, Double value1, Double value2, Double value3) {
            this.name = name;
            this.category=category;
            this.value1 = value1;
            this.value2 = value2;
            this.value3 = value3;
        }

        // Getter 和 Setter
        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public Double getValue1() {
            return value1;
        }

        public void setValue1(Double value1) {
            this.value1 = value1;
        }

        public Double getValue2() {
            return value2;
        }

        public void setValue2(Double value2) {
            this.value2 = value2;
        }

        public Double getValue3() {
            return value3;
        }

        public void setValue3(Double value3) {
            this.value3 = value3;
        }
    }


    // 示例数据
    private static List<Data> generateData() {
        List<Data> dataList = new ArrayList<>();
        dataList.add(new Data("A","北京", 10.5, 20.0, 30.0));
        dataList.add(new Data("B", "北京",15.0, 25.0, 35.0));
        dataList.add(new Data("C","上海", 20.0, 30.0, 40.0));
        return dataList;
    }


    /**
     * 自动计算数值类型字段的合计值，支持任意类型 T。
     *
     * @param data 数据列表
     * @param <T>  数据类型
     * @return 含有合计值的对象
     */
    public static  <T> T calculateTotalRow(List<T> data) {
        if (data == null || data.isEmpty()) {
            throw new IllegalArgumentException("数据列表不能为空");
        }

        // 创建合计行对象
        T totalRow;
        try {
            totalRow = (T) data.get(0).getClass().getDeclaredConstructor().newInstance();
        } catch (Exception e) {
            throw new RuntimeException("无法创建合计行对象", e);
        }

        // 遍历字段计算合计
        for (Field field : data.get(0).getClass().getDeclaredFields()) {
            field.setAccessible(true);
            if (Number.class.isAssignableFrom(field.getType())) {
                try {
                    // 计算当前字段的合计值
                    double sum = data.stream()
                            .mapToDouble(item -> {
                                try {
                                    Object value = field.get(item);
                                    return value == null ? 0 : ((Number) value).doubleValue();
                                } catch (IllegalAccessException e) {
                                    throw new RuntimeException(e);
                                }
                            })
                            .sum();

                    // 根据字段类型赋值
                    if (field.getType() == Integer.class) {
                        field.set(totalRow, (int) sum);
                    } else if (field.getType() == Double.class) {
                        field.set(totalRow, sum);
                    } else if (field.getType() == Long.class) {
                        field.set(totalRow, (long) sum);
                    } else if (field.getType() == BigDecimal.class) {
                        field.set(totalRow, new BigDecimal(sum) );
                    }
                } catch (IllegalAccessException e) {
                    throw new RuntimeException(e);
                }
            } else if (field.getType() == String.class) {
                try {
                    // 设置合计行的标志字段（如 "合计"）
                    if ("name".equalsIgnoreCase(field.getName())) {
                        field.set(totalRow, "合计");
                    }
                } catch (IllegalAccessException e) {
                    throw new RuntimeException(e);
                }
            }
        }
        return totalRow;
    }


    /**
     * 头部样式
     */
    private static WriteCellStyle getHeadStyle() {
        // 头的策略
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        // 背景颜色
        headWriteCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headWriteCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);

        // 字体
        WriteFont headWriteFont = new WriteFont();
        //设置字体名字
        headWriteFont.setFontName("微软雅黑");
        //设置字体大小
        headWriteFont.setFontHeightInPoints((short) 10);
        //字体加粗
        headWriteFont.setBold(false);
        //在样式用应用设置的字体
        headWriteCellStyle.setWriteFont(headWriteFont);

        // 边框样式
        setBorderStyle(headWriteCellStyle);

        //设置自动换行
        headWriteCellStyle.setWrapped(true);

        //设置水平对齐的样式为居中对齐
        headWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        //设置垂直对齐的样式为居中对齐
        headWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        //设置文本收缩至合适
        //        headWriteCellStyle.setShrinkToFit(true);

        return headWriteCellStyle;
    }


    /**
     * 内容样式
     */
    private static WriteCellStyle getContentStyle() {
        // 内容的策略
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();

        // 背景白色
        // 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND 不然无法显示背景颜色.头默认了 FillPatternType所以可以不指定
        contentWriteCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        contentWriteCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);

        // 设置字体
        WriteFont contentWriteFont = new WriteFont();
        //设置字体大小
        contentWriteFont.setFontHeightInPoints((short) 10);
        //设置字体名字
        contentWriteFont.setFontName("宋体");
        //在样式用应用设置的字体
        contentWriteCellStyle.setWriteFont(contentWriteFont);

        //设置样式
        setBorderStyle(contentWriteCellStyle);

        // 水平居中
        contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        // 垂直居中
        contentWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置自动换行
        contentWriteCellStyle.setWrapped(true);

        //设置单元格格式是：文本格式，方式长数字文本科学计数法
//        contentWriteCellStyle.setDataFormatData();

        //设置文本收缩至合适
        // contentWriteCellStyle.setShrinkToFit(true);

        return contentWriteCellStyle;
    }

    /**
     * 边框样式
     */
    private static void setBorderStyle(WriteCellStyle cellStyle) {
        //设置底边框
        cellStyle.setBorderBottom(BorderStyle.THIN);
        //设置底边框颜色
        cellStyle.setBottomBorderColor(IndexedColors.BLACK1.getIndex());
        //设置左边框
        cellStyle.setBorderLeft(BorderStyle.THIN);
        //设置左边框颜色
        cellStyle.setLeftBorderColor(IndexedColors.BLACK1.getIndex());
        //设置右边框
        cellStyle.setBorderRight(BorderStyle.THIN);
        //设置右边框颜色
        cellStyle.setRightBorderColor(IndexedColors.BLACK1.getIndex());
        //设置顶边框
        cellStyle.setBorderTop(BorderStyle.THIN);
        //设置顶边框颜色
        cellStyle.setTopBorderColor(IndexedColors.BLACK1.getIndex());
    }

}
