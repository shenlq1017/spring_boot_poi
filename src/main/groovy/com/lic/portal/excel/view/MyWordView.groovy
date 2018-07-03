package com.lic.portal.excel.view

import cn.afterturn.easypoi.word.WordExportUtil
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.springframework.stereotype.Controller
import org.springframework.web.servlet.view.AbstractView

import javax.servlet.ServletOutputStream
import javax.servlet.http.HttpServletRequest
import javax.servlet.http.HttpServletResponse

@Controller
class MyWordView extends MyMapBaseView {
    private static final String CONTENT_TYPE = "application/msword";

    public MyWordView() {
        setContentType(CONTENT_TYPE);
    }

    @Override
    protected void renderMergedOutputModel(Map<String, Object> model, HttpServletRequest request,
                                           HttpServletResponse response) throws Exception {
        String codedFileName = "临时文件.docx";
        if (model.containsKey("fileName")) {
            codedFileName = (String) model.get("fileName") + ".doc";
        }
        if (isIE(request)) {
            codedFileName = java.net.URLEncoder.encode(codedFileName, "UTF8");
        } else {
            codedFileName = new String(codedFileName.getBytes("UTF-8"), "ISO-8859-1");
        }
        response.setHeader("content-disposition", "attachment;filename=" + codedFileName);
        XWPFDocument document = WordExportUtil.exportWord07(
                (String) model.get("url"),
                (Map<String, Object>) model.get("map"));
        ServletOutputStream out = response.getOutputStream();
        document.write(out);
        out.flush();
    }
}
