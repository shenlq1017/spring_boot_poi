package com.lic.portal.excel.util

import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity
import org.apache.poi.ss.usermodel.Sheet

class MyCellUtil {

    private static final int CELL_WIDTH = 10;

    /**
     * 初始化默认全部都是10
     * @param sheet
     * @param cloumn
     */
    static def initCellWidth(Sheet sheet, int cloumnNums) {
        for (int i = 0; i < cloumnNums; i++) {
            sheet.setColumnWidth(i, (int) (256 * 10));
        }
    }

    /**
     * 设置某列宽度
     * @param excelParam
     * @param sheet
     * @param cellx
     * @return
     */
    static def cellWidth(ExcelExportEntity excelParam, Sheet sheet,int cellx) {
        if(excelParam.getWidth() != CELL_WIDTH) {
            sheet.setColumnWidth(cellx, (int) (256 * excelParam.getWidth()));
        }
    }
}
