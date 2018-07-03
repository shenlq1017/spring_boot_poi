package com.lic.portal.excel.util

import org.apache.poi.hssf.record.cf.FontFormatting
import org.apache.poi.hssf.usermodel.HSSFCellStyle
import org.apache.poi.hssf.usermodel.HSSFFont
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.RegionUtil

class MyCellStyleUtil {

    static CellStyle onBorder(CellStyle cellStyle,short borderInt) {
        cellStyle.setBorderBottom(borderInt);
        cellStyle.setBorderLeft(borderInt)
        cellStyle.setBorderRight(borderInt)
        cellStyle.setBorderTop(borderInt)
        return cellStyle;
    }

    /**
     * 给合并单元格设置边框
     * @param i
     * @param cellRangeTitle
     * @param sheet
     * @param workBook
     */
    static void onBorderForMergeCell(int i, CellRangeAddress cellRangeTitle, Sheet sheet, Workbook workBook){
        RegionUtil.setBorderBottom(i, cellRangeTitle, sheet, workBook);
        RegionUtil.setBorderLeft(i, cellRangeTitle, sheet, workBook);
        RegionUtil.setBorderRight(i, cellRangeTitle, sheet, workBook);
        RegionUtil.setBorderTop(i, cellRangeTitle, sheet, workBook);
    }



    static CellStyle onFontBold(Workbook workbook,CellStyle cellStyle,short fontHeight) {
        HSSFFont font =workbook.createFont();
        font.setBold(true);
        font.setFontHeight((short) (fontHeight*20))
        cellStyle.setFont(font);
        return cellStyle
    }

    static CellStyle onFontUnderLineAndBold(Workbook workbook,CellStyle cellStyle,short fontHeight,boolean underLine,boolean isBold) {
        HSSFFont font = workbook.createFont();
        if (underLine) {
            font.setUnderline(FontFormatting.U_SINGLE);
        }
        if (isBold) {
            font.setBold(true);
        }
        font.setFontHeight((short) (fontHeight*20))
        cellStyle.setFont(font);
        return cellStyle
    }
}
