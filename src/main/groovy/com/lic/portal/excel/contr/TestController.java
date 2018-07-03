package com.lic.portal.excel.contr;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

/**
 * description
 *
 * @author sxx
 * @date 2018/07/02 10:48
 */
@Controller
public class TestController {

    @RequestMapping("upload")
    public String upload() {
        return "upload";
    }
}
