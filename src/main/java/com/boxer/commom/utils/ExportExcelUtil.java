package com.boxer.commom.utils;

import com.alibaba.excel.EasyExcel;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.List;

@Slf4j
public class ExportExcelUtil {
    public static <T> void export(HttpServletResponse response, List<T> records, String fileName,Class<?> clazz) {
        log.info("Exporting Excel file: {}", fileName);
        try (OutputStream output=response.getOutputStream()){
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setCharacterEncoding(StandardCharsets.UTF_8.name());
            String encodedFileName = URLEncoder.encode(fileName, StandardCharsets.UTF_8);
            response.setHeader("Content-Disposition", "attachment;filename*=UTF-8''" + encodedFileName);
            EasyExcel.write(output, clazz).sheet("Sheet1").doWrite(records);
            log.info("Excel export completed successfully.");
        } catch (IOException e) {
            throw new RuntimeException("Excel export failed due to an I/O error.", e);
        }
    }
}
