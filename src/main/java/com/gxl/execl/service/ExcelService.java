package com.gxl.execl.service;

import com.gxl.execl.bean.Excel;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.text.ParseException;
import java.util.List;

public interface ExcelService {
    public List<Excel> dealExcel(String filePath) throws IOException, ParseException;
}
