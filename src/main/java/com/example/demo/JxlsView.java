package com.example.demo;

import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.Map;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.web.servlet.view.AbstractView;

public class JxlsView extends AbstractView {
    private String templatePath;
    private String exportFileName;

    public JxlsView(String templatePath, String exportFileName) {
        this.templatePath = templatePath;
        if (exportFileName != null) {
            try {
                exportFileName = URLEncoder.encode(exportFileName, "UTF-8");
            } catch (UnsupportedEncodingException e) {
                e.printStackTrace();
            }
        }
        this.exportFileName = exportFileName;
        setContentType("application/vnd.ms-excel");
    }

    @Override
    protected void renderMergedOutputModel(Map<String, Object> map, HttpServletRequest request, HttpServletResponse response) throws Exception {
        try (InputStream is = getClass().getClassLoader().getResourceAsStream(templatePath)) {
            response.setContentType(getContentType());
            response.setHeader("content-disposition", "attachment;filename=" + exportFileName + ".xls");
            ServletOutputStream os = response.getOutputStream();

            JxlsUtils.exportExcel(templatePath, os, map);
            
            os.close();
        }
    }
    
}
