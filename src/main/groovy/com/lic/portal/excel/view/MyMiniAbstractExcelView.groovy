package com.lic.portal.excel.view

import cn.afterturn.easypoi.view.PoiBaseView

abstract class MyMiniAbstractExcelView extends MyMapBaseView {

    private static final String CONTENT_TYPE = "text/html;application/vnd.ms-excel";

    protected static final String HSSF = ".xls";
    protected static final String XSSF = ".xlsx";

    public MiniAbstractExcelView() {
        setContentType(CONTENT_TYPE);
    }

}