package com.gxl.execl.controller;

import com.gxl.execl.bean.Excel;
import com.gxl.execl.service.ExcelService;
import com.gxl.execl.util.CSVRead;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
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
    @RequestMapping("/upload")
    public String  uploadExcel(@RequestParam("file") MultipartFile image,Model model) throws Exception {
//        List<Excel> excels = excelService.dealExcel(realPath);
        String name = image.getOriginalFilename();
        File file =multipartFileToFile(image);
        CSVRead.CSVDelete(file);
//        model.addAttribute("quanChengs",excels);
//        return "allrecord";
        return "redirect:/show?realPath="+"src/main/java/com/gxl/execl/util/new.xls";
    }

    private  File multipartFileToFile(MultipartFile file) throws Exception {
        File toFile = null;
        if (file.equals("") || file.getSize() <= 0) {
            file = null;
        } else {
            InputStream ins = null;
            ins = file.getInputStream();
            toFile = new File(file.getOriginalFilename());
            inputStreamToFile(ins, toFile);
            ins.close();
        }
        return toFile;

    }
    private static void inputStreamToFile(InputStream ins, File file) {
        try {
            OutputStream os = new FileOutputStream(file);
            int bytesRead = 0;
            byte[] buffer = new byte[8192];
            while ((bytesRead = ins.read(buffer, 0, 8192)) != -1) {
                os.write(buffer, 0, bytesRead);
            }
            os.close();
            ins.close();
        } catch (Exception e) {
                e.printStackTrace();
        }
    }


}
