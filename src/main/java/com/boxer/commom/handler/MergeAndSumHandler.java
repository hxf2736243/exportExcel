//package com.boxer.commom.handler;
//import com.alibaba.excel.write.handler.WriteHandler;
//import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.ss.util.CellRangeAddress;
//
//public class MergeAndSumHandler implements WriteHandler {
//    private final int columnToMerge; // 需要合并的列数
//    private final String sumLabel;  // 合计标签
//    private final int sumColumnIndex; // 合计值所在的列
//
//    public MergeAndSumHandler(int columnToMerge, String sumLabel, int sumColumnIndex) {
//        this.columnToMerge = columnToMerge;
//        this.sumLabel = sumLabel;
//        this.sumColumnIndex = sumColumnIndex;
//    }
//
//    public void afterSheetCreate(WriteSheetHolder writeSheetHolder) {
//        Sheet sheet = writeSheetHolder.getSheet();
//        Workbook workbook = sheet.getWorkbook();
//        int lastRowIndex = sheet.getLastRowNum(); // 获取最后一行的索引
//
//        // 动态计算合计值（示例：计算指定列的总和）
//        double sum = 0;
//        for (int i = 0; i <= lastRowIndex; i++) {
//            Row row = sheet.getRow(i);
//            if (row != null) {
//                Cell cell = row.getCell(sumColumnIndex); // 获取指定列
//                if (cell != null && cell.getCellType() == CellType.NUMERIC) {
//                    sum += cell.getNumericCellValue();
//                }
//            }
//        }
//
//        // 创建合计行
//        Row sumRow = sheet.createRow(lastRowIndex + 1);
//        Cell firstCell = sumRow.createCell(0);
//        firstCell.setCellValue(sumLabel);
//
//        // 设置样式
//        CellStyle centerStyle = workbook.createCellStyle();
//        centerStyle.setAlignment(HorizontalAlignment.CENTER);
//        centerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
//        firstCell.setCellStyle(centerStyle);
//
//        // 合并单元格（例如从第 0 列到第 columnToMerge - 1 列）
//        CellRangeAddress range = new CellRangeAddress(lastRowIndex + 1, lastRowIndex + 1, 0, columnToMerge - 1);
//        sheet.addMergedRegion(range);
//
//        // 设置合计值
//        Cell sumCell = sumRow.createCell(sumColumnIndex);
//        sumCell.setCellValue(sum);
//        sumCell.setCellStyle(centerStyle);
//    }
//}
