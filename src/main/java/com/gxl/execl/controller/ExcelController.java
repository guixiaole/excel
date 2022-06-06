package com.gxl.execl.controller;

import com.gxl.execl.bean.Excel;
import com.gxl.execl.service.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;

import java.io.IOException;
import java.text.ParseException;
import java.util.List;

@Controller
public class ExcelController {
    @Autowired
    private ExcelService excelService;
    @RequestMapping("/show")
    public String  showExcel(String realPath, Model model) throws IOException, ParseException {
        List<Excel> excels = excelService.dealExcel(realPath);
        model.addAttribute("quanChengs",excels);
        return "allrecord";
    }

}
