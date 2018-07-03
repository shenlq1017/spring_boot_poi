package com.lic.portal.excel.model

class CellMerge {

    static final String MERVE_COLUMN_KEY = "mergeKey";

    String cellKey;

    String cellVal;

    int mergeLengthways = 0 ;

    int mergeCrosswise = 0;

    CellMerge() {
    }

    CellMerge(String cellKey, String cellVal, int mergeLengthways, int mergeCrosswise) {
        this.cellKey = cellKey
        this.cellVal = cellVal
        this.mergeLengthways = mergeLengthways
        this.mergeCrosswise = mergeCrosswise
    }
}
