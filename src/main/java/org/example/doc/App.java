package org.example.doc;

import org.apache.commons.lang3.tuple.Pair;

import java.util.List;

/**
 * Hello world!
 */
public class App {

    // 1 立项决议
//    static String template = "前5个文档/1、立项决议w/立项决议基础模板.docx";
//    static String dataExecl = "前5个文档/1、立项决议w/立项决议替换数据表.xlsx";
//    static String target = "前5个文档/1、立项决议w/result/%s_立项决议.docx";
//    static String sheetName = "";

    // 2、会议纪要w
//    static String template = "前5个文档/2、会议纪要w/会议纪要基础模板.docx";
//    static String dataExecl = "前5个文档/2、会议纪要w/会议纪要替换数据表.xlsx";
//    static String target = "前5个文档/2、会议纪要w/result/%s_会议纪要.docx";
//    static String sheetName = "";

    // 3、验收意见表w
    static String template = "前5个文档/3、验收意见表w/验收意见表基础模板.docx";
    static String dataExcel = "前5个文档/3、验收意见表w/验收意见表替换数据表.xlsx";
    static String target = "前5个文档/3、验收意见表w/result/%s_验收意见表.docx";
    static String sheetName = "";


    public static void main(String[] args) throws Exception {
        // 数据来源文档路径
        List<List<Pair<String,String>>> projectData = DocUtil.readExcel(dataExcel,sheetName);

        // 生成文件路径  %s 替换为项目名写死在第一列
        String format = target;

        DocUtil.replaceDoc(template, format, projectData);
    }



}