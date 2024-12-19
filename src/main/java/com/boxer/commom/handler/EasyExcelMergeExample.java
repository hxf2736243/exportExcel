//package com.boxer.commom.handler;
//
//import com.alibaba.excel.EasyExcel;
//
//import java.util.ArrayList;
//import java.util.List;
//
//import com.alibaba.excel.metadata.Head;
//import com.alibaba.excel.metadata.data.WriteCellData;
//import com.alibaba.excel.write.handler.CellWriteHandler;
//import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
//import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
//import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.ss.util.CellRangeAddress;
//
//
//public class EasyExcelMergeExample {
//
//    public static void main(String[] args) throws Exception {
//        // 创建示例数据
//        List<List<String>> data = new ArrayList<>();
//        data.add(List.of("A", "B", "C"));
//        data.add(List.of("A", "B", "D"));
//        data.add(List.of("A", "B", "D"));
//        data.add(List.of("A", "B", "E"));
////
////        // 输出文件路径
////        String fileName = "merged_cells_example_with_cellwritehandler.xlsx";
////
////        // 使用 EasyExcel 写入文件，并注册自定义的 CellWriteHandler 来处理合并单元格
////        EasyExcel.write(fileName)
////                .sheet("Sheet1")
////                .registerWriteHandler(new CellWriteHandler() {
////                    @Override
////                    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<WriteCellData<?>> list, Cell cell,
////                                                 Head head,
////                                                 Integer integer, Boolean aBoolean) {
////                        Sheet sheet = context.getWriteSheetHolder().getSheet();
////                        int totalRows = sheet.getPhysicalNumberOfRows();
////
////                        if (totalRows == 0) return;
////
////                        Row firstRow = sheet.getRow(0);
////                        if (firstRow == null) return;
////
////                        int totalCols = firstRow.getPhysicalNumberOfCells();
////
////                        // 合并相邻行相同列的单元格
////                        int startRow = 0;
////                        for (int row = 1; row <= totalRows; row++) {
////                            Row currentRow = sheet.getRow(row);
////                            Row previousRow = sheet.getRow(row - 1);
////
////                            // 如果当前行或上一行是 null，跳过该行
////                            if (currentRow == null || previousRow == null) {
////                                continue;
////                            }
////
////                            int col = 0;
////                            String currentValue = currentRow.getCell(col) != null ? currentRow.getCell(col).getStringCellValue() : "";
////                            String previousValue = previousRow.getCell(col) != null ? previousRow.getCell(col).getStringCellValue() : "";
////
////                            // 当前单元格与前一个单元格的值不同，进行合并
////                            if (currentValue.equals(previousValue)) {
////                                // 如果有多行，进行合并
////                                if (row - startRow >= 1) {
////                                    // 合并的区域：startRow 到 row-1，列索引 col
////                                    sheet.addMergedRegion(new CellRangeAddress(startRow, row - 1, col, col));
////                                }
////                                // 更新起始行
////                                startRow = row;
////                            }
////                        }
////                    }
////                })
////                .doWrite(data);
//    }
//}
//
