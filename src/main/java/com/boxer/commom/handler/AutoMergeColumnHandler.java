//package com.boxer.commom.handler;
//import com.alibaba.excel.write.handler.RowWriteHandler;
//import com.alibaba.excel.write.handler.context.RowWriteHandlerContext;
//import com.alibaba.excel.write.handler.context.SheetWriteHandlerContext;
//import lombok.extern.slf4j.Slf4j;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.util.CellRangeAddress;
//
//import java.util.ArrayList;
//import java.util.List;
//
//
//@Slf4j
//public class AutoMergeColumnHandler implements RowWriteHandler {
//    private final int columnIndex; // 需要合并的列索引，从 0 开始
//    private List<String> columnValues; // 存储每一行的数据，用于合并判断
//
//    public AutoMergeColumnHandler(int columnIndex) {
//        this.columnIndex = columnIndex;
//    }
//
//    @Override
//    public void afterRowCreate(RowWriteHandlerContext context) {
//        // 每一行数据写入后，判断是否需要合并
//        Sheet sheet = context.getWriteSheetHolder().getSheet();
//        Row row = context.getRow();
//
//        // 获取当前行列值
//        String currentValue = row.getCell(columnIndex).toString();
//
//        // 存储合并操作的逻辑
//        if (columnValues == null) {
//            columnValues = new ArrayList<>();
//        }
//
//        // 判断是否需要合并
//        int currentRowIndex = row.getRowNum();
//        if (currentRowIndex > 0 && !columnValues.get(currentRowIndex - 1).equals(currentValue)) {
//            // 当前行的值与上一行不同，进行合并
//            sheet.addMergedRegion(new CellRangeAddress(currentRowIndex - 1, currentRowIndex - 1, columnIndex, columnIndex));
//        }
//
//        columnValues.add(currentValue);
//    }
//}
