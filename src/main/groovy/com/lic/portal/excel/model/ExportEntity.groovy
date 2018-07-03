package com.lic.portal.excel.model

import cn.afterturn.easypoi.excel.entity.ExportParams
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity

class ExportEntity {
    ExportParamsHeader entity;
    List<ExcelExportEntity> entityList;
    Collection<?> dataSet;
    ExportParamsFoot exportParamsFoot;


    ExportEntity() {
    }

    ExportEntity(ExportParamsHeader entity, List<ExcelExportEntity> entityList, Collection<?> dataSet) {
        this.entity = entity
        this.entityList = entityList
        this.dataSet = dataSet
    }
}
