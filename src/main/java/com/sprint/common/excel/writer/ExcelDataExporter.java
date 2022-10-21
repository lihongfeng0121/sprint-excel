package com.sprint.common.excel.writer;

import com.sprint.common.excel.util.Excels;
import com.sprint.common.excel.util.FileUtils;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.nio.file.Files;

/**
 * @author hongfeng.li
 * @since 2022/10/20
 */
public class ExcelDataExporter {


    private final ExcelWriter excelWriter;

    public ExcelDataExporter(ExcelWriter excelWriter) {
        this.excelWriter = excelWriter;
    }

    public void export(OutputStream outputStream) throws IOException {
        excelWriter.export(outputStream);
    }

    public void export(String filePath) throws IOException {
        if (filePath.endsWith(excelWriter.getFileName())) {
            filePath = filePath.replace(excelWriter.getFileName(), "");
        }

        FileUtils.createFolder(filePath);

        export(Files.newOutputStream(new File(FileUtils.resolvePath(filePath, excelWriter.getFileName())).toPath()));
    }

    public void export(javax.servlet.http.HttpServletResponse response) throws IOException {
        response.setContentType(Excels.getContentType(excelWriter.getWorkbook()));
        try {
            response.setHeader("Content-Disposition",
                    "attachment; filename=" + URLEncoder.encode(excelWriter.getFileName(), "UTF-8"));
        } catch (UnsupportedEncodingException var3) {
            response.setHeader("Content-Disposition", "attachment; filename=" + excelWriter.getFileName());
        }

        export(response.getOutputStream());
    }
}
