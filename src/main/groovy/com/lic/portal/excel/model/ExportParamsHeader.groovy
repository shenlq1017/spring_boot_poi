package com.lic.portal.excel.model

import cn.afterturn.easypoi.excel.entity.ExportParams

class ExportParamsHeader extends ExportParams {

    boolean needUnderLine = true;

    boolean needBold = true;

    short wordSize = (short) 25;

    ExportParamsHeader(String title, String secondTitle, String sheetName) {
        super(title, secondTitle, sheetName)
    }
}
