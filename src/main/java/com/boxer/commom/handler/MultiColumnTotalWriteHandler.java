//package com.boxer.commom.handler;
//
//import com.alibaba.excel.write.handler.WriteHandler;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.ss.util.CellRangeAddress;
//
//import java.util.Map;
//
//public class MultiColumnTotalWriteHandler implements WriteHandler {
//    private final int totalRowIndex; // 合计行的行号
//    private final String totalText; // 合计文本
//    private final Map<Integer, Double> columnSums; // 每列的合计值
//    private final int mergeStartColumn; // 合并单元格的起始列
//    private final int mergeEndColumn;   // 合并单元格的结束列
//
//    public MultiColumnTotalWriteHandler(int totalRowIndex, String totalText, Map<Integer, Double> columnSums, int mergeStartColumn, int mergeEndColumn) {
//        this.totalRowIndex = totalRowIndex;
//        this.totalText = totalText;
//        this.columnSums = columnSums;
//        this.mergeStartColumn = mergeStartColumn;
//        this.mergeEndColumn = mergeEndColumn;
//    }
//
//    public void sheet(int sheetNo, Sheet sheet) {
//        // 合并单元格
//        if (mergeStartColumn < mergeEndColumn) {
//            sheet.addMergedRegion(new CellRangeAddress(totalRowIndex, totalRowIndex, mergeStartColumn, mergeEndColumn));
//        }
//    }
//
//    public void row(int rowIndex, Row row) {
//        // 不处理
//    }
//
//    public void cell(int cellIndex, Cell cell) {
//        if (cell.getRowIndex() == totalRowIndex) {
//            CellStyle cellStyle = cell.getSheet().getWorkbook().createCellStyle();
//            Font font = cell.getSheet().getWorkbook().createFont();
//            font.setBold(true);
//            cellStyle.setFont(font);
//            cell.setCellStyle(cellStyle);
//
//            if (cell.getColumnIndex() == mergeStartColumn) {
//                cell.setCellValue(totalText); // 合并单元格区域中的第一列写入合计文本
//            } else if (cell.getColumnIndex() > mergeStartColumn && cell.getColumnIndex() <= mergeEndColumn) {
//                cell.setCellValue(""); // 合并单元格区域中的其他列清空
//            } else {
//                Double totalValue = columnSums.get(cell.getColumnIndex());
//                if (totalValue != null) {
//                    cell.setCellValue(totalValue);
//                }
//            }
//        }
//    }
//}
