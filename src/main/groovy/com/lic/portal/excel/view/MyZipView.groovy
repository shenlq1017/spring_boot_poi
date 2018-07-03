package com.lic.portal.excel.view

import cn.afterturn.easypoi.word.WordExportUtil
import com.lic.portal.excel.model.ExportEntity
import com.lic.portal.excel.util.MyExportUtil
import com.lic.portal.excel.view.bean.MapExcelManyConstants
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xwpf.usermodel.XWPFDocument

import javax.servlet.ServletOutputStream
import javax.servlet.http.HttpServletRequest
import javax.servlet.http.HttpServletResponse
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream

class MyZipView  extends MyMapBaseView {

    //application/octet-stream
    private static final String CONTENT_TYPE = "application/octet-stream";

    @Override
    protected void renderMergedOutputModel(Map<String, Object> model, HttpServletRequest request,
                                           HttpServletResponse response) throws Exception {
        String codedFileName = "临时文件";
        Workbook workbook = MyExportUtil.exportExcel((List<ExportEntity>) model.get(MapExcelManyConstants.MANY_EXPORT_ENTITY),
                model.get(MapExcelManyConstants.EXCEL_TYPE),model.get(MapExcelManyConstants.DATA_SIZE));
        if (model.containsKey(MapExcelManyConstants.FILE_NAME)) {
            codedFileName = (String) model.get(MapExcelManyConstants.FILE_NAME);
        }
        if (workbook instanceof HSSFWorkbook) {
            codedFileName += HSSF;
        } else {
            codedFileName += XSSF;
        }
        if (isIE(request)) {
            codedFileName = java.net.URLEncoder.encode(codedFileName, "UTF8");
        } else {
            codedFileName = new String(codedFileName.getBytes("UTF-8"), "ISO-8859-1");
        }
        response.setHeader("content-disposition", "attachment;filename=" + codedFileName);
        ServletOutputStream out = response.getOutputStream();
        workbook.write(out);
        out.flush();
    }

    static Workbook renderManyOutPut(Map<String, Object> model) {
        Workbook workbook = MyExportUtil.exportExcel((List<ExportEntity>) model.get(MapExcelManyConstants.MANY_EXPORT_ENTITY),
                model.get(MapExcelManyConstants.EXCEL_TYPE),model.get(MapExcelManyConstants.DATA_SIZE));
        return workbook;
    }

    static def outZip(List<Workbook> workbookList,List<String> fileNames,String zipName, HttpServletRequest request,
               HttpServletResponse response) {
        OutputStream out = null;
        // 设置导出excel文件
        out = response.getOutputStream();
        ZipOutputStream zipOutputStream = new ZipOutputStream(out);
        zipName = "批量文件" + ".zip";
        response.setContentType(CONTENT_TYPE);
        response.setHeader("Connection", "close"); // 表示不能用浏览器直接打开
        response.setHeader("Accept-Ranges", "bytes");// 告诉客户端允许断点续传多线程连接下载
        response.setHeader("Content-Disposition",
                "attachment;filename=" + new String(zipName.getBytes("GB2312"), "ISO8859-1"));

        for (int i = 0; i < workbookList.size(); i++) {
            ZipEntry entry = new ZipEntry(fileNames.get(i) + ".xls");
            zipOutputStream.putNextEntry(entry);
            Workbook workbook = workbookList.get(i);
            workbook.write(zipOutputStream);
        }
        // 关闭输出流
        zipOutputStream.flush();
        zipOutputStream.close();
    }


}
