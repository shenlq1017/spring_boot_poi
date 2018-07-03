package com.lic.portal.excel.model

class CellMapping {

    int cellIndex = 0;

    int DataType =1;

    CellMapping() {
    }

    CellMapping(int cellIndex, int dataType) {
        this.cellIndex = cellIndex
        DataType = dataType
    }
}
