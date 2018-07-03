package com.lic.portal.excel.service

import cn.afterturn.easypoi.excel.entity.TemplateExportParams
import com.lic.portal.excel.view.MyMapBaseView
import org.springframework.stereotype.Service
import org.springframework.ui.ModelMap

import javax.servlet.http.HttpServletRequest
import javax.servlet.http.HttpServletResponse

@Service
class WordTemplateOutService {


    def downloadByPoiBaseView(ModelMap modelMap, HttpServletRequest request,
                                      HttpServletResponse response) {
        Map<String, Object> map = new HashMap<String, Object>();


        map.put("company","单位A")
        map.put("payYear","2018")
        map.put("payMonth","06")
        map.put("moneyc","127388")
        map.put("moneyh","1273")
        map.put("exchangeRate","10.001")
        map.put("askYear","2018")
        map.put("aksMonth","06")
        map.put("aksDay","30")

        modelMap.put("fileName", "申请");
        modelMap.put("map", map);
        modelMap.put("url","D:/data/portal/excel/src/main/resources/doc/test.docx")
        MyMapBaseView.render(modelMap, request, response,"templateWord");

    }
}
