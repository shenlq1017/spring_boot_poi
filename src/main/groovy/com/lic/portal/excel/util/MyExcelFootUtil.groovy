package com.lic.portal.excel.util

import com.lic.portal.excel.model.ExportParamsFoot
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellRangeAddress

class MyExcelFootUtil {

    static void creatFoot(int index, Sheet sheet, Workbook workbook,ExportParamsFoot exportParamsFoot,int fieldLength) {
        int rowIndex = index+2;
        Row row = sheet.createRow(rowIndex);
        row.setHeight((short) (20*20));
        CellStyle footStyle = workbook.createCellStyle();
        footStyle = MyCellStyleUtil.onFontUnderLineAndBold(workbook,footStyle,exportParamsFoot.getWordSize(),exportParamsFoot.getNeedUnderLine(),exportParamsFoot.isNeedBold());
        footStyle.setAlignment(CellStyle.ALIGN_CENTER);

        String footStr="";
        List<String> stringList = exportParamsFoot.getFootStrs()
        String spiltWord = exportParamsFoot.getSpiltWord();
        for (int i = 0; i < stringList.size(); i++) {
            if(i==stringList.size()-1) {
                spiltWord = "";
            }
            footStr+=stringList.get(i)+spiltWord;
        }
//        String footStr= "22222";//exportParamsFoot.toString();
        new MyExportUtil().createStringCell(row, 0, footStr, footStyle, null);
        CellRangeAddress cellRangeAddressFoot = new CellRangeAddress(rowIndex, rowIndex, 0, fieldLength-1);
        sheet.addMergedRegion(cellRangeAddressFoot);
        MyCellStyleUtil.onBorderForMergeCell(0,cellRangeAddressFoot,sheet,workbook);
    }
}
