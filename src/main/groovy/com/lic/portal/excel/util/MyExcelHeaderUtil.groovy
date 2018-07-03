package com.lic.portal.excel.util

import com.lic.portal.excel.model.ExportParamsHeader
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

class MyExcelHeaderUtil {
    static CellStyle setTitleStyle(Sheet sheet, Workbook workbook, ExportParamsHeader exportParamsHeader, CellStyle cellStyle) {
        cellStyle = MyCellStyleUtil.onFontUnderLineAndBold(workbook,cellStyle,exportParamsHeader.getWordSize(),exportParamsHeader.getNeedUnderLine(),exportParamsHeader.isNeedBold());
        return cellStyle;
    }
}
