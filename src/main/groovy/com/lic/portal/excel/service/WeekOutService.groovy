package com.lic.portal.excel.service

import cn.afterturn.easypoi.entity.vo.MapExcelConstants
import cn.afterturn.easypoi.excel.entity.ExportParams
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity
import cn.afterturn.easypoi.excel.entity.vo.BaseEntityTypeConstants
import com.fasterxml.jackson.annotation.JsonInclude
import com.fasterxml.jackson.databind.ObjectMapper
import com.lic.portal.excel.model.ExportEntity
import com.lic.portal.excel.model.ExportParamsHeader
import com.lic.portal.excel.view.MyMapBaseView
import com.lic.portal.excel.view.MyZipView
import com.lic.portal.excel.view.bean.MapExcelManyConstants
import net.sf.json.JSONObject
import org.apache.poi.ss.usermodel.Workbook
import org.springframework.stereotype.Service
import org.springframework.ui.ModelMap

import javax.servlet.http.HttpServletRequest
import javax.servlet.http.HttpServletResponse

@Service
class WeekOutService {


    def weekout(ModelMap modelMap, HttpServletRequest request,
                HttpServletResponse response) {
        List<ExcelExportEntity> entity = new ArrayList<ExcelExportEntity>();
        ExcelExportEntity excelentityNum = new ExcelExportEntity("编号", "idNum");
        excelentityNum.setNeedMerge(true);
        excelentityNum.setMergeVertical(true);
        excelentityNum.setMergeRely(2)
        ExcelExportEntity excelentityName = new ExcelExportEntity("品名", "name");
//        excelentityName.setMergeVertical(true);
//        excelentityName.setNeedMerge(true)
        ExcelExportEntity excelentityUnit = new ExcelExportEntity("单位", "unit");
//        excelentityUnit.setNeedMerge(true)
//        excelentityNum.setMergeVertical(true);
        ExcelExportEntity excelentity = new ExcelExportEntity(null, "company");
        List<ExcelExportEntity> temp = new ArrayList<ExcelExportEntity>();
        temp.add(new ExcelExportEntity("名称", "cname"));
        temp.add(new ExcelExportEntity("简称", "csimpleName"));
        excelentity.setList(temp)

        entity.add(excelentityNum)
        entity.add(excelentityName)
        entity.add(excelentityUnit)
        entity.add(excelentity)

        modelMap.put(MapExcelConstants.MAP_LIST, new ArrayList<Map<String, Object>>());
        modelMap.put(MapExcelConstants.ENTITY_LIST, entity);
        ExportParams params = new ExportParams("主副食品一周计划采购统计清单", "公司一", ExcelType.XSSF);
        modelMap.put(MapExcelConstants.PARAMS, params)
        modelMap.put(MapExcelConstants.FILE_NAME, "主副食品一周计划");
        downloadManyView(modelMap, request, response)
    }


    def downloadManyView(ModelMap modelMap, HttpServletRequest request,
                           HttpServletResponse response) {
        modelMap.put(MapExcelManyConstants.DATA_SIZE,10000)
        modelMap.put(MapExcelManyConstants.EXCEL_TYPE,ExcelType.HSSF)

        List<ExportEntity> exportEntities = new ArrayList<>()


        for (int i = 0; i < 2; i++) {
            List<Map<String,Object>> datas = new ArrayList<>()

            for(int m =0;m<300;m++) {
                Map<String,Object> map = new HashMap<>()
                map.put("idNum",m);
                map.put("oName","商品"+m);
                map.put("priceUnit","kg")
                map.put("price",23.23)
                map.put("total",30231.23)

                for (int l = 0; l < 6; l++) {
                    for (int k = 0; k < 7; k++) {
                        map.put("weekCountChild3"+l+k,Math.random())
                    }
                }
//                datas.add(map)
            }

            ExportEntity exportEntity = new ExportEntity();
            List<ExcelExportEntity> entity = new ArrayList<ExcelExportEntity>();
            ExcelExportEntity excelentityNum = new ExcelExportEntity("编号", "idNum");
            excelentityNum.setWidth(6)
            excelentityNum.setNeedMerge(true);
            excelentityNum.setMergeRely(0, 2);

            ExcelExportEntity excelentityObjName = new ExcelExportEntity("品名", "oName");
            excelentityObjName.setWidth(16)
            excelentityObjName.setNeedMerge(true);
            excelentityObjName.setMergeRely(0, 2);

            ExcelExportEntity priceUnit = new ExcelExportEntity("单位", "priceUnit");
            priceUnit.setWidth(4)
            priceUnit.setNeedMerge(true);
            priceUnit.setMergeRely(0, 2);

            ExcelExportEntity price = new ExcelExportEntity("价格", "price");
            price.setType(BaseEntityTypeConstants.DOUBLE_TYPE)
            price.setNeedMerge(true);
            price.setMergeRely(0, 2);

            int companyNums = 2;
            entity.add(excelentityNum)
            entity.add(excelentityObjName)
            entity.add(priceUnit)
            entity.add(price)
            for (int l = 0; l < companyNums; l++) {
                ExcelExportEntity excelentityName = new ExcelExportEntity(null, "name");
                ExcelExportEntity excelentityNum2 = new ExcelExportEntity("名称"+l, "idNum2"+l);
                excelentityNum2.setWidth(5)
                excelentityNum2.setNeedMerge(true);
                excelentityNum2.setMergeRely(6, 0);
                ExcelExportEntity excelentityNum3 = new ExcelExportEntity("S9#-"+l, "idNum3"+l);
                excelentityNum3.setNeedMerge(true);
                excelentityNum3.setMergeRely(6, 0);
                List<ExcelExportEntity> exportEntitiesTemp = [excelentityNum2, excelentityNum3];
                excelentityName.setList(exportEntitiesTemp);


                entity.add(excelentityName)

                for (int k = 0; k < 7; k++) {
                    ExcelExportEntity weekCount = new ExcelExportEntity(null, "weekCount"+l+k);
                    ExcelExportEntity weekChild = new ExcelExportEntity(null, "weekCountChild1"+l+k);
                    ExcelExportEntity weekChild2 = new ExcelExportEntity(null, "weekCountChild2"+l+k);
                    ExcelExportEntity weekChild3 = new ExcelExportEntity(intToStr(k+1), "weekCountChild3"+l+k);
                    weekChild3.setType(BaseEntityTypeConstants.DOUBLE_TYPE)
                    weekChild3.setWidth(5)
                    weekCount.setList([weekChild, weekChild2, weekChild3]);
                    entity.add(weekCount);
                }
            }

            ExcelExportEntity total = new ExcelExportEntity("合计", "total");
            total.setType(BaseEntityTypeConstants.DOUBLE_TYPE)
            total.setNeedMerge(true);
            total.setMergeRely(0,2);


            entity.add(total)

            exportEntity.setEntityList(entity)

            ExportParamsHeader exportEntityTitle = new ExportParamsHeader("主副食品一周计划采购统计清单"+i, "制表单位:工作单位一&&制表单位:工作单位二","公司一"+i);
//            exportEntityTitle.setHeight((short)20)
            exportEntity.setEntity(exportEntityTitle)
            exportEntity.setDataSet(datas)
            exportEntities.add(exportEntity)
        }

        modelMap.put(MapExcelManyConstants.MANY_EXPORT_ENTITY,exportEntities)
        modelMap.put(MapExcelManyConstants.FILE_NAME,"一周计划")



        //单个导出
//        MyMapBaseView.render(modelMap, request, response, MapExcelManyConstants.MY_MAP_EXCEL_VIEW);

        //zip导出
        List<Workbook> workbookList = new ArrayList<>();
        List<String> fileNames = new ArrayList<>();
        for (int i = 0; i < 20; i++) {
            Workbook workbook = MyZipView.renderManyOutPut(modelMap);
            workbookList.add(workbook);
            fileNames.add("一周计划"+i);
        }
        MyZipView.outZip(workbookList,fileNames,"out.zip",request,response);
    }

    String intToStr(int i) {
        switch (i){
            case 1:
                return "一";
            case 2:
                return "二";
            case 3:
                return "三";
            case 4:
                return "四";
            case 5:
                return "五";
            case 6:
                return "六";
            case 7:
                return "日";
        }
    }

}
