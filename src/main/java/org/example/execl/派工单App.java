package org.example.execl;

import java.util.List;

/**
 * 派工单Excel
 */
public class 派工单App {

    public static final String a1 = "2020年%s自主研究开发项目组派工单";
    public static final String a2 = "项目名称：%s           编号：%s";
    public static String sheetName2 = "研发审计项目总表数据";

    public static String filePath = "6个QA文档/替换数据汇总表.xlsx";
    public static String templatePath = "6个QA文档/派工单基础模板.xlsx";
    public static String targetFilePath = "6个QA文档/result派工单/%s_派工单.xlsx";

    public static String sheetName = "工时分配表-蜂助手";

    public static void main(String[] args) throws Exception {

        List<ExcelUtil.Project> projects = ExcelUtil.readExcel(filePath, sheetName, true);
        for (ExcelUtil.Project project : projects) {
            ExcelUtil.writeExcel(targetFilePath, project);
        }
    }




}
