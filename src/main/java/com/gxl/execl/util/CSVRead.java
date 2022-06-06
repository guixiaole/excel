package com.gxl.execl.util;

import com.gxl.execl.bean.Excel;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;

public class CSVRead {
    public static void CSVDelete( File fileDes)throws IOException, ParseException{
//        File fileDes = new File(realPath);
//        InputStream str = new FileInputStream(fileDes);
//        Workbook book = new HSSFWorkbook(str);
        Workbook book = null;
        try {
            book = new XSSFWorkbook(new FileInputStream(fileDes));
        } catch (Exception ex) {
            book = new HSSFWorkbook(new FileInputStream(fileDes));
        }
//        XSSFWorkbook book = new XSSFWorkbook(str);
        Sheet sheet =  book.getSheetAt(0);
        int rows=sheet.getLastRowNum();//总行数
        int i = 0;
        List<Excel> excels = new ArrayList<>();
        int firstStart = 0;
        int flag =0;

        for (;i<=rows;i++){
//            if (rows-i<4){
//                break;
//            }
//            System.out.println(i);
            String event = String.valueOf(sheet.getRow(i).getCell(1));

            System.out.println(event);
            if(event.equals("过信号机")||event.equals("进站")||event.equals("出站") ){

                firstStart=1;
                flag++;
                Excel excel = new Excel(String.valueOf(sheet.getRow(i).getCell(0)),
                        String.valueOf(sheet.getRow(i).getCell(1)),
                        String.valueOf(sheet.getRow(i).getCell(2)),
                        String.valueOf(sheet.getRow(i).getCell(3)),
                        String.valueOf(sheet.getRow(i).getCell(4)),
                        String.valueOf(sheet.getRow(i).getCell(5)),
                        String.valueOf(sheet.getRow(i).getCell(6)),
                        String.valueOf(sheet.getRow(i).getCell(7)),
                        String.valueOf(sheet.getRow(i).getCell(8)),
                        String.valueOf(sheet.getRow(i).getCell(9)),
                        String.valueOf(sheet.getRow(i).getCell(10)),
                        String.valueOf(sheet.getRow(i).getCell(11)),
                        String.valueOf(sheet.getRow(i).getCell(12)),
                        String.valueOf(sheet.getRow(i).getCell(13)),
                        String.valueOf(sheet.getRow(i).getCell(14)),
                        String.valueOf(sheet.getRow(i).getCell(15)),
                        String.valueOf(sheet.getRow(i).getCell(16)),
                        String.valueOf(sheet.getRow(i).getCell(17))
                );
//                System.out.println(excel);
                excels.add(excel);
            }
            else if (firstStart ==0){
                Excel excel = new Excel(String.valueOf(sheet.getRow(i).getCell(0)),
                        String.valueOf(sheet.getRow(i).getCell(1)),
                        String.valueOf(sheet.getRow(i).getCell(2)),
                        String.valueOf(sheet.getRow(i).getCell(3)),
                        String.valueOf(sheet.getRow(i).getCell(4)),
                        String.valueOf(sheet.getRow(i).getCell(5)),
                        String.valueOf(sheet.getRow(i).getCell(6)),
                        String.valueOf(sheet.getRow(i).getCell(7)),
                        String.valueOf(sheet.getRow(i).getCell(8)),
                        String.valueOf(sheet.getRow(i).getCell(9)),
                        String.valueOf(sheet.getRow(i).getCell(10)),
                        String.valueOf(sheet.getRow(i).getCell(11)),
                        String.valueOf(sheet.getRow(i).getCell(12)),
                        String.valueOf(sheet.getRow(i).getCell(13)),
                        String.valueOf(sheet.getRow(i).getCell(14)),
                        String.valueOf(sheet.getRow(i).getCell(15)),
                        String.valueOf(sheet.getRow(i).getCell(16)),
                        String.valueOf(sheet.getRow(i).getCell(17))
                );
//                System.out.println(excel);
                excels.add(excel);
            }else{
//                continue;
                sheet.removeRow(sheet.getRow(i));
//                if (rows-i>4){
//                    sheet.shiftRows(i,rows,-1);
//                }else {
//                    break;
//                }

//                int flagAll = sheet.getLastRowNum();
//                System.out.println(flagAll);
//                i--;
            }
        }

        Workbook wb = new HSSFWorkbook();
        HSSFSheet sheet1 = (HSSFSheet) wb.createSheet();
        flag = 0;
        int flagJ= 0;
        for (int j = 0;j<excels.size();j++){

            HSSFRow row = sheet1.createRow(flagJ);
            flagJ++;
            if (excels.get(j).getEvent().equals("进站")||excels.get(j).getEvent().equals("出站")){
                flag = 1;
            }

            row.createCell(0).setCellValue(excels.get(j).getXuHao());
            row.createCell(1).setCellValue(excels.get(j).getEvent());

            row.createCell(2).setCellValue(excels.get(j).getDateTime()==null?" ":excels.get(j).getDateTime());
            row.createCell(3).setCellValue(excels.get(j).getGongLiBiao()==null?" ":excels.get(j).getGongLiBiao());
            row.createCell(4).setCellValue(excels.get(j).getOther()==null?" ":excels.get(j).getOther());
            row.createCell(5).setCellValue(excels.get(j).getDistance()==null?" ":excels.get(j).getDistance());
            row.createCell(6).setCellValue(excels.get(j).getSignalMachine()==null?" ":excels.get(j).getSignalMachine());
            row.createCell(7).setCellValue(excels.get(j).getXinHao()==null?" ":excels.get(j).getXinHao());
            row.createCell(8).setCellValue(excels.get(j).getSpeed()==null?" ":excels.get(j).getSpeed());
            row.createCell(9).setCellValue(excels.get(j).getRestrictSpeed()==null?" ":excels.get(j).getRestrictSpeed());
            row.createCell(10).setCellValue(excels.get(j).getLingWei()==null?" ":excels.get(j).getLingWei());
            row.createCell(11).setCellValue(excels.get(j).getQianYin()==null?" ":excels.get(j).getQianYin());
            row.createCell(12).setCellValue(excels.get(j).getQianHou()==null?" ":excels.get(j).getQianHou());
            row.createCell(13).setCellValue(excels.get(j).getGuanYa()==null?" ":excels.get(j).getGuanYa());
            row.createCell(14).setCellValue(excels.get(j).getGangYa()==null?" ":excels.get(j).getGangYa());
            row.createCell(15).setCellValue(excels.get(j).getZhuanSuDianLiu()==null?" ":excels.get(j).getZhuanSuDianLiu());
            row.createCell(16).setCellValue(excels.get(j).getJunGang1()==null?" ":excels.get(j).getJunGang1());
            row.createCell(17).setCellValue(excels.get(j).getJunGang2()==null?" ":excels.get(j).getJunGang2());
            if (flag==1){
                HSSFRow row1 = sheet1.createRow(flagJ);
                flagJ++;
                HSSFRow row2 = sheet1.createRow(flagJ);
                flagJ++;
                HSSFRow row3 = sheet1.createRow(flagJ);
                flagJ++;
            }
        }
        wb.setSheetName(0,"修改后全程记录");
        FileOutputStream fileOut = new FileOutputStream("src/main/java/com/gxl/execl/util/new.xls");
        wb.write(fileOut);
        fileOut.close();

        //        System.out.println(excels);

    }

    public static List<Excel> CsvRead(String realPath)throws IOException, ParseException{
        File fileDes = new File(realPath);
//        InputStream str = new FileInputStream(fileDes);
//        Workbook book = new HSSFWorkbook(str);
        Workbook book = null;
        try {
            book = new XSSFWorkbook(new FileInputStream(fileDes));
        } catch (Exception ex) {
            book = new HSSFWorkbook(new FileInputStream(fileDes));
        }
//        XSSFWorkbook book = new XSSFWorkbook(str);
        Sheet sheet =  book.getSheetAt(0);
        int rows=sheet.getLastRowNum();//总行数
        List<Excel> excels = new ArrayList<>();
        for (int i =1;i<rows;i++){
                Excel excel = new Excel(String.valueOf(sheet.getRow(i).getCell(0)),
                        String.valueOf(sheet.getRow(i).getCell(1)),
                        String.valueOf(sheet.getRow(i).getCell(2)),
                        String.valueOf(sheet.getRow(i).getCell(3)),
                        String.valueOf(sheet.getRow(i).getCell(4)),
                        String.valueOf(sheet.getRow(i).getCell(5)),
                        String.valueOf(sheet.getRow(i).getCell(6)),
                        String.valueOf(sheet.getRow(i).getCell(7)),
                        String.valueOf(sheet.getRow(i).getCell(8)),
                        String.valueOf(sheet.getRow(i).getCell(9)),
                        String.valueOf(sheet.getRow(i).getCell(10)),
                        String.valueOf(sheet.getRow(i).getCell(11)),
                        String.valueOf(sheet.getRow(i).getCell(12)),
                        String.valueOf(sheet.getRow(i).getCell(13)),
                        String.valueOf(sheet.getRow(i).getCell(14)),
                        String.valueOf(sheet.getRow(i).getCell(15)),
                        String.valueOf(sheet.getRow(i).getCell(16)),
                        String.valueOf(sheet.getRow(i).getCell(17))
                );
                excels.add(excel);
        }
        return excels;
    }
    public static void main(String[] args) throws IOException, ParseException {
        String realPath = "src/main/java/com/gxl/execl/util/K769-7120704-7120256K0.0411F.xlsx";
//        CSVRead.CSVDelete(realPath);
    }
}
