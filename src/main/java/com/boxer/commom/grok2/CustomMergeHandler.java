package com.boxer.commom.grok2;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

public class CustomMergeHandler implements CellWriteHandler {
    private final List<?> dataList;           // 导出数据
    private final Set<Integer> globalMergeCols;  // 全局合并的列索引集合
    private final Set<Integer> groupMergeCols;   // 分组合并的列索引集合
    private final int groupByCol;             // 分组依据的列索引

    private final int headHeight;
    private final int summaryMergeStartCol; // 总计行合并起始列
    private final int summaryMergeEndCol;   // 总计行合并结束列


    public CustomMergeHandler(List<?> dataList,int headHeight, int[] globalMergeCols, int[] groupMergeCols, int groupByCol,int summaryMergeStart,int summaryMergeEnd) {
        this.dataList = dataList;
        this.headHeight=headHeight;
        this.globalMergeCols = new HashSet<>();
        this.groupMergeCols = new HashSet<>();
        for (int col : globalMergeCols) {
            this.globalMergeCols.add(col);
        }
        for (int col : groupMergeCols) {
            this.groupMergeCols.add(col);
        }
        this.groupByCol = groupByCol;
        this.summaryMergeStartCol=summaryMergeStart;
        this.summaryMergeEndCol=summaryMergeEnd;
    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {

        if (isHead || dataList.isEmpty()) {
            return; // 跳过头部的处理或空数据
        }

        int colIndex = cell.getColumnIndex();
        int rowIndex = cell.getRowIndex() - headHeight; // 减去头部行

        if (rowIndex < 0 || rowIndex >= dataList.size()) {
            return;
        }

        Object currentData = dataList.get(rowIndex);
        boolean isSummaryRow = "总计".equals(getCellValue(currentData, 0)); // 检查是否为汇总行（姓名列为“总计”）

        if (isSummaryRow && colIndex == 0) {
            writeSheetHolder.getSheet().addMergedRegion(new CellRangeAddress(rowIndex + headHeight, rowIndex + headHeight,summaryMergeStartCol , summaryMergeEndCol));
        } else if (globalMergeCols.contains(colIndex)) {
            mergeGlobal(writeSheetHolder.getSheet(), rowIndex, colIndex);
        } else if (groupMergeCols.contains(colIndex)) {
            mergeByGroup(writeSheetHolder.getSheet(), rowIndex, colIndex);
        }
    }

    // 全局合并逻辑
    private void mergeGlobal(Sheet sheet, int rowIndex, int colIndex) {
        if (rowIndex > 0) {
            Object prevData = dataList.get(rowIndex - 1);
            Object currentData = dataList.get(rowIndex);
            if (getCellValue(prevData, colIndex).equals(getCellValue(currentData, colIndex))) {
                return; // 如果与前一行值相同，跳过（避免重复合并）
            }
        }

        int startRow = rowIndex;
        int endRow = rowIndex;
        Object value = getCellValue(dataList.get(rowIndex), colIndex);

        for (int i = rowIndex + 1; i < dataList.size(); i++) {
            Object nextData = dataList.get(i);
            Object nextValue = getCellValue(nextData, colIndex);
            if (value.equals(nextValue)) {
                endRow = i;
            } else {
                break;
            }
        }

        if (endRow > startRow) {
            sheet.addMergedRegion(new CellRangeAddress(startRow + headHeight, endRow + headHeight, colIndex, colIndex));
        }
    }

    // 按组合并逻辑
    private void mergeByGroup(Sheet sheet, int rowIndex, int colIndex) {
        if (rowIndex > 0) {
            Object prevData = dataList.get(rowIndex - 1);
            Object currentData = dataList.get(rowIndex);
            Object prevGroupValue = getCellValue(prevData, groupByCol);
            Object currentGroupValue = getCellValue(currentData, groupByCol);
            Object prevValue = getCellValue(prevData, colIndex);
            Object currentValue = getCellValue(currentData, colIndex);
            if (currentValue.equals(prevValue) && currentGroupValue.equals(prevGroupValue)) {
                return; // 如果与前一行值相同且分组相同，跳过
            }
        }

        int startRow = rowIndex;
        int endRow = rowIndex;
        Object currentData = dataList.get(rowIndex);
        Object value = getCellValue(currentData, colIndex);
        Object groupValue = getCellValue(currentData, groupByCol);

        for (int i = rowIndex + 1; i < dataList.size(); i++) {
            Object nextData = dataList.get(i);
            Object nextGroupValue = getCellValue(nextData, groupByCol);
            Object nextValue = getCellValue(nextData, colIndex);

            if (groupValue.equals(nextGroupValue) && value.equals(nextValue)) {
                endRow = i;
            } else {
                break;
            }
        }

        if (endRow > startRow) {
            sheet.addMergedRegion(new CellRangeAddress(startRow + headHeight, endRow + headHeight, colIndex, colIndex));
        }
    }

    // 获取单元格值（假设数据是 DTO，通过反射获取字段值）
    private Object getCellValue(Object obj, int colIndex) {
        try {
            Field[] fields = obj.getClass().getDeclaredFields();
            if (colIndex >= 0 && colIndex < fields.length) {
                Field field = fields[colIndex];
                field.setAccessible(true);
                return field.get(obj);
            }
            return null;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
}