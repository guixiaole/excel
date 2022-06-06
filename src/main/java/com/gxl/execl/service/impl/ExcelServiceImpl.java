package com.gxl.execl.service.impl;

import com.gxl.execl.bean.Excel;
import com.gxl.execl.service.ExcelService;
import com.gxl.execl.util.CSVRead;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.text.ParseException;
import java.util.List;

@Service
public class ExcelServiceImpl implements ExcelService {

    @Override
    public List<Excel> dealExcel(String filePath) throws IOException, ParseException {
        return CSVRead.CsvRead(filePath);
    }
}
