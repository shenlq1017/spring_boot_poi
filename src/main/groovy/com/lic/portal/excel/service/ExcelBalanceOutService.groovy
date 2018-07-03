package com.lic.portal.excel.service

import cn.afterturn.easypoi.excel.entity.ExportParams
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity
import cn.afterturn.easypoi.excel.entity.vo.BaseEntityTypeConstants
import com.lic.portal.excel.model.CellMerge
import com.lic.portal.excel.model.ExportEntity
import com.lic.portal.excel.model.ExportParamsFoot
import com.lic.portal.excel.model.ExportParamsHeader
import com.lic.portal.excel.view.MyMapBaseView
import com.lic.portal.excel.view.bean.MapExcelManyConstants
import org.apache.poi.ss.usermodel.CellStyle
import org.springframework.stereotype.Service
import org.springframework.ui.ModelMap

import javax.servlet.http.HttpServletRequest
import javax.servlet.http.HttpServletResponse
@Service
class ExcelBalanceOutService {


    def downloadManyView(ModelMap modelMap, HttpServletRequest request,
                         HttpServletResponse response) {
        modelMap.put(MapExcelManyConstants.DATA_SIZE,10000)
        modelMap.put(MapExcelManyConstants.EXCEL_TYPE,ExcelType.HSSF)
        ExportEntity exportEntity = new ExportEntity();
        ExportParamsHeader exportEntityTitle = new ExportParamsHeader("XX单位申请代付2018年**月份主副食供应结算明细表", "制表单位:工作单位一&&时间： 2018年1月5日","公司一");
        List<ExcelExportEntity> entity = new ArrayList<ExcelExportEntity>();
        ExcelExportEntity area = new ExcelExportEntity("区域","area");
        area.setWidth(8);
        ExcelExportEntity idNum = new ExcelExportEntity("序号","idNum");
        idNum.setWidth(6);
        ExcelExportEntity company = new ExcelExportEntity("公司","company");
        company.setWrap(true)
        company.setWidth(30);
        ExcelExportEntity moneyc = new ExcelExportEntity("人民币（元）","moneyc");
        moneyc.setType(BaseEntityTypeConstants.DOUBLE_TYPE)
        moneyc.setWidth(16);
        ExcelExportEntity moneyh = new ExcelExportEntity("港币（元）","moneyh");
        moneyh.setType(BaseEntityTypeConstants.DOUBLE_TYPE)
        moneyh.setWidth(16);
        ExcelExportEntity memo = new ExcelExportEntity("备注","memo");
        memo.setWidth(14);

        entity = [area,idNum,company,moneyc,moneyh,memo];

        ExportParamsFoot exportParamsFoot = new ExportParamsFoot();
        List<String> foots = ["申请人：张胜","审批人：李四","财务：王五","审批单位：大东海"]
        exportParamsFoot.setFootStrs(foots);
        exportParamsFoot.setSpiltWord("             ")

        exportEntity.setEntityList(entity);

        //数据
        List<Map<String, Object>> datas = new ArrayList<>()
        for(int j =0;j<2;j++) {
            String areaName= "北京";
            if(j==1) {
                areaName = "上海";
            }
            for (int i = 0; i < 15; i++) {
                Map<String, Object> map = new HashMap<>();
                map.put("area", areaName)
                map.put("idNum", i + 1);
                map.put("company", "这是一个名字比较长比较长比较长的单位" + i);
                map.put("moneyc", 200031.42)
                map.put("moneyh", 20031.42)
                map.put("memo", "备注")
                datas.add(map);
            }
            datas.get(j*15).put(CellMerge.MERVE_COLUMN_KEY,[new CellMerge("area",areaName,14,0)]);
        }
        Map<String,Object> total = new HashMap<>()
        total.put("moneyc",200031.42*30);
        total.put("moneyh",20031.42*30);
        total.put(CellMerge.MERVE_COLUMN_KEY,[new CellMerge("area","合计",0,2)])
        datas.add(total)


        exportEntity.setDataSet(datas)

        exportEntityTitle.setNeedUnderLine(false)
        exportEntityTitle.setWordSize((short)16)
        exportEntity.setEntity(exportEntityTitle);
        exportEntity.setExportParamsFoot(exportParamsFoot)

        modelMap.put(MapExcelManyConstants.MANY_EXPORT_ENTITY,[exportEntity])
        modelMap.put(MapExcelManyConstants.FILE_NAME,"结算明细")
        MyMapBaseView.render(modelMap, request, response, MapExcelManyConstants.MY_MAP_EXCEL_VIEW);
    }
}
