package org.example.execl;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.sun.org.apache.regexp.internal.RE;
import org.apache.commons.collections4.map.ListOrderedMap;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

public class 任务分解表App {

    public static String filePath = "6个QA文档/替换数据汇总表.xlsx";
    public static String sheetName = "工时分配表-蜂助手";
    public static String sheetName2 = "研发审计项目总表数据";

    public static String templatePath = "6个QA文档/6.任务分解表基础模板.xlsx";
    public static String targetFilePath = "6个QA文档/result任务分解表/%s_任务分解表.xlsx";

    public static void main(String[] args) throws Exception {
        List<ExcelUtil.Project> projects = ExcelUtil.readExcel(filePath, sheetName, false);
        Map<String, Map<String, String>> projectData = ExcelUtil.getProjectData(filePath, sheetName2);
        writeFile(projects, projectData);
    }

    private static void writeFile(List<ExcelUtil.Project> projects, Map<String, Map<String, String>> projectData) throws Exception{
        for (ExcelUtil.Project project : projects) {
            Map<String, String> map = projectData.get(project.getName());
            String fileName = String.format(targetFilePath, project.name);
            FileInputStream inputStream = new FileInputStream(templatePath);
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

            XSSFSheet sheet1 = workbook.getSheetAt(0);
            XSSFSheet sheet2 = workbook.getSheetAt(1);
            XSSFSheet sheet3 = workbook.getSheetAt(2);
            sheet1(project.getName(), map, sheet1);
            sheet2(project.getName(), map, sheet2);
            sheet3(project, map, sheet3);

            workbook.write(new FileOutputStream(new File(fileName)));
            workbook.close();
        }
    }

    private static void sheet1(String projectName, Map<String, String> map, XSSFSheet sheet1) {
        if(map == null){
            map = new HashMap<>();
        }
        String fm1 = "项目名称：%s        负责人：%s                     编号：%s";
        String s1 = String.format(fm1, projectName, map.getOrDefault("项目负责人",""),
                map.getOrDefault("项目编号", ""));
        sheet1.getRow(1).getCell(0).setCellValue(s1);
        String fm2 = "项目立项：%s";
        String s2 = String.format(fm2, map.getOrDefault("立项时间",""));
        sheet1.getRow(3).getCell(4).setCellValue(s2);
        String fm3 = "进入开发：%s";
        String s3 = String.format(fm3, map.getOrDefault("进入开发时间",""));
        sheet1.getRow(4).getCell(4).setCellValue(s3);
        String fm4 = "项目结束：%s";
        String s4 = String.format(fm4, map.getOrDefault("完成开发时间",""));
        sheet1.getRow(12).getCell(4).setCellValue(s4);
    }

    private static void sheet2(String projectName, Map<String, String> map, XSSFSheet sheet) {
        if(map == null){
            map = new HashMap<>();
        }
        String fm1 = "项目名称：%s        负责人：%s                     编号：%s";
        String s1 = String.format(fm1, projectName, map.getOrDefault("项目负责人",""),
                map.getOrDefault("项目编号", ""));
        sheet.getRow(1).getCell(0).setCellValue(s1);
    }

    private static void sheet3(ExcelUtil.Project project, Map<String, String> map, XSSFSheet sheet) {
        if(map == null){
            map = new HashMap<>();
        }
        String fm1 = "项目名称：%s        负责人：%s                     编号：%s";
        String s1 = String.format(fm1, project.getName(), map.getOrDefault("项目负责人",""),
                map.getOrDefault("项目编号", ""));
        sheet.getRow(1).getCell(0).setCellValue(s1);

        LinkedHashMap<String, List<ExcelUtil.YueShuJu>> yueShuJu = project.yueShuJu;

        Map<ExcelUtil.YueShuJu,List<Double>> result = new LinkedHashMap<>();
        for (Map.Entry<String, List<ExcelUtil.YueShuJu>> entry : yueShuJu.entrySet()) {
            List<ExcelUtil.YueShuJu> value = entry.getValue();
            for (ExcelUtil.YueShuJu shuJu : value) {
                List<Double> list = result.getOrDefault(shuJu, new ArrayList<>(12));
                list.add(shuJu.getNumber());
                result.putIfAbsent(shuJu, list);
            }
        }

        int startRow = 4;
        int i = 0;
        for (Map.Entry<ExcelUtil.YueShuJu, List<Double>> yueShuJuListEntry : result.entrySet()) {
            ExcelUtil.YueShuJu key = yueShuJuListEntry.getKey();
            List<Double> value = yueShuJuListEntry.getValue();
            Double aDouble = yueShuJuListEntry.getValue().stream().reduce((x, y) -> x + y).get();
            if(aDouble <= 0) continue;
            XSSFRow row = sheet.createRow(startRow + i);
            row.createCell(0).setCellValue(i+1);
            row.createCell(1).setCellValue(key.get项目角色());
            row.createCell(2).setCellValue(key.getName());
            for (int j = 0; j <= 12; j++) {
                if(j == 12){
                    row.createCell(3+j).setCellValue(aDouble);
                }else if( j < value.size()){
                    row.createCell(3+j).setCellValue(value.get(j));
                }else{
                    row.createCell(3+j).setCellValue(0);
                }
            }
            i++;
        }

        CellRangeAddress region = new CellRangeAddress(startRow + i,startRow + i,0,2);
        sheet.addMergedRegion(region);
        XSSFRow row = sheet.createRow(startRow + i);
        row.createCell(0).setCellValue("汇总");

        for (int j = 0; j <= 12; j++) {
            double sum = 0.0;
            for (int k = 0; k < i; k++) {
                XSSFRow row1 = sheet.getRow(startRow + k);
                double numericCellValue = row1.getCell(j + 3).getNumericCellValue();
                sum = sum + numericCellValue;
            }
            row.createCell(j+3).setCellValue(sum);
        }
    }
}
