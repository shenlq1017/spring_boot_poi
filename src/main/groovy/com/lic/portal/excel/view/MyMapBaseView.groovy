package com.lic.portal.excel.view

import cn.afterturn.easypoi.entity.vo.*
import cn.afterturn.easypoi.view.*
import com.lic.portal.excel.view.bean.MapExcelManyConstants
import org.slf4j.Logger
import org.slf4j.LoggerFactory
import org.springframework.web.servlet.view.AbstractView

import javax.servlet.http.HttpServletRequest
import javax.servlet.http.HttpServletResponse

abstract class MyMapBaseView extends AbstractView {

    private static final Logger LOGGER = LoggerFactory.getLogger(MyMapBaseView.class);

    protected static boolean isIE(HttpServletRequest request) {
        return (request.getHeader("USER-AGENT").toLowerCase().indexOf("msie") > 0
                || request.getHeader("USER-AGENT").toLowerCase().indexOf("rv:11.0") > 0
                || request.getHeader("USER-AGENT").toLowerCase().indexOf("edge") > 0) ? true
                : false;
    }

    static void render(Map<String, Object> model, HttpServletRequest request,
                              HttpServletResponse response, String viewName) {
        MyMapBaseView view = null;
        if (BigExcelConstants.EASYPOI_BIG_EXCEL_VIEW.equals(viewName)) {
            view = new EasypoiBigExcelExportView();
        } else if (MapExcelConstants.EASYPOI_MAP_EXCEL_VIEW.equals(viewName)) {
            view = new EasypoiMapExcelView();
        } else if (NormalExcelConstants.EASYPOI_EXCEL_VIEW.equals(viewName)) {
            view = new EasypoiSingleExcelView();
        } else if (TemplateExcelConstants.EASYPOI_TEMPLATE_EXCEL_VIEW.equals(viewName)) {
            view = new EasypoiTemplateExcelView();
        } else if (MapExcelGraphConstants.MAP_GRAPH_EXCEL_VIEW.equals(viewName)) {
            view = new MapGraphExcelView();
        } else if (MapExcelManyConstants.MY_MAP_EXCEL_VIEW.equals(viewName)) {
            view = new MyMapView();
        } else if ("templateWord".equals(viewName)) {
            view = new MyWordView();
        }
        try {
            view.renderMergedOutputModel(model, request, response);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        }
    }

}