package com.lic.portal.excel.contr

import com.lic.portal.excel.service.ExcelBalanceOutService
import com.lic.portal.excel.service.WeekOutService
import com.lic.portal.excel.service.WordTemplateOutService
import com.lic.portal.excel.util.MyImportUtil
import org.springframework.beans.factory.annotation.Autowired
import org.springframework.ui.ModelMap
import org.springframework.web.bind.annotation.GetMapping
import org.springframework.web.bind.annotation.PostMapping
import org.springframework.web.bind.annotation.RequestMapping
import org.springframework.web.bind.annotation.RequestParam
import org.springframework.web.bind.annotation.RestController
import org.springframework.web.multipart.MultipartFile

import javax.servlet.http.HttpServletRequest
import javax.servlet.http.HttpServletResponse

@RestController
@RequestMapping("/outweek")
class WeekOutContr {

    @Autowired
    WeekOutService weekOutService;


    @Autowired
    WordTemplateOutService wordTemplateOutService;

    @Autowired
    ExcelBalanceOutService excelBalanceOutService;

    @Autowired
    MyImportUtil myImportUtil;

    @GetMapping
    def weekout(ModelMap modelMap, HttpServletRequest request,
                HttpServletResponse response) {
        weekOutService.downloadManyView(modelMap, request, response)
    }

    @GetMapping("balanceout")
    def balanceout(ModelMap modelMap, HttpServletRequest request,
                HttpServletResponse response) {
        excelBalanceOutService.downloadManyView(modelMap, request, response)
    }

    @GetMapping("/downword")
    def wordout(ModelMap modelMap, HttpServletRequest request,
                HttpServletResponse response) {
        wordTemplateOutService.downloadByPoiBaseView(modelMap,request,response)
    }

    @PostMapping("/upload")
    def importExcel(@RequestParam("file") MultipartFile file) {
        myImportUtil.importByInput(file);
    }
}
