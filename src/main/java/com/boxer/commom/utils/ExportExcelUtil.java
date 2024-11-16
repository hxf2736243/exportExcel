package com.boxer.commom.utils;

import com.alibaba.excel.EasyExcel;
import jakarta.servlet.http.HttpServletResponse;

import java.io.IOException;
import java.net.URLEncoder;
import java.util.List;

public class ExportExcelUtil {
    public static <T> void export(HttpServletResponse response, List<T> records, String fileName,Class<?> clazz)throws IOException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        String encodedFileName = URLEncoder.encode(fileName, "UTF-8");
        response.setHeader("Content-Disposition", "attachment;filename*=UTF-8''" + encodedFileName);
        EasyExcel.write(response.getOutputStream(), clazz).sheet("Sheet1").doWrite(records);
    }
}
