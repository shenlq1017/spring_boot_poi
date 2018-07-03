package com.lic.portal.excel.util

import cn.afterturn.easypoi.excel.ExcelImportUtil
import cn.afterturn.easypoi.excel.entity.ImportParams
import org.apache.commons.io.FileUtils
import org.springframework.stereotype.Service
import org.springframework.web.multipart.MultipartFile


@Service
class MyImportUtil {

    List<Map<String,Object>> importByInput(MultipartFile file) {
        ImportParams params = new ImportParams();
        params.setTitleRows(1);
        params.setHeadRows(1);
//        params.setDataHandler(new ExcelDataHandlerDefaultImpl());

        File fileOnly = new File("/test/"+file.getOriginalFilename())
        FileUtils.copyInputStreamToFile(file.getInputStream(), fileOnly);
//        file.transferTo(fileOnly)
        InputStream is = new FileInputStream(fileOnly);
        List<Map<String, Object>> list = ExcelImportUtil.importExcel(is, Map.class, params);
        fileOnly.deleteOnExit()
        for (int i = 0; i < list.size(); i++) {
            for(Map.Entry<String,Object> map :list.get(i).entrySet()) {
                println("key-"+map.getKey()+":val-"+map.getValue());
            }
        }
    }
}
