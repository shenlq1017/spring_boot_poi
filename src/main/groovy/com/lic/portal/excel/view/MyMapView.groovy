package com.lic.portal.excel.view

import cn.afterturn.easypoi.entity.vo.MapExcelConstants
import cn.afterturn.easypoi.entity.vo.MapExcelGraphConstants
import cn.afterturn.easypoi.excel.ExcelExportUtil
import cn.afterturn.easypoi.excel.entity.ExportParams
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity
import cn.afterturn.easypoi.view.MiniAbstractExcelView
import com.lic.portal.excel.model.ExportEntity
import com.lic.portal.excel.util.MyExportUtil
import com.lic.portal.excel.view.bean.MapExcelManyConstants
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import org.springframework.stereotype.Controller

import javax.servlet.ServletOutputStream
import javax.servlet.http.HttpServletRequest
import javax.servlet.http.HttpServletResponse


@Controller(MapExcelManyConstants.MY_MAP_EXCEL_VIEW)
class MyMapView extends MyMiniAbstractExcelView {

    MyMapView() {
    }

    /**
     * List<ExportEntity> exportEntities, ExcelType excelType,int dataSize
     * @param model
     * @param request
     * @param response
     * @throws Exception
     */
    @Override
    protected void renderMergedOutputModel(Map<String, Object> model, HttpServletRequest request, HttpServletResponse response) throws Exception {
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
}
