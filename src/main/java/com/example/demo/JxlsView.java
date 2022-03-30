package com.example.demo;

import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.Map;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
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
            response.setHeader("content-disposition", "attachment;filename=" + exportFileName + ".xlsx");
            ServletOutputStream os = response.getOutputStream();

            
            // test/1
            Context content = new Context();
            content.putVar("user", map);

            // test/2
            // Context content = PoiTransformer.createInitialContext();
            // content.putVar("list", map);

            JxlsHelper.getInstance().processTemplate(is, os, content);
            
            os.close();
        }
    }
    
}
