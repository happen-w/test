package org.example.execl;

import com.alibaba.fastjson2.JSON;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

public class 工时审核表App {

    static String a1 = "%s年%s自主研究开发项目组人员工时审批表";
    static String a2 = "项目名称：%s                                 负责人：%s                                   编号：%s";

    static String filePath = "6个QA文档/工时审核-每月/202102工时.xlsx";
    static String sheetName = "Sheet1";
    static String templatePath = "6个QA文档/5.研发工时审核表基础模板.xlsx";
    static String targetFilePath = "6个QA文档/result工时审核表/%s_工时审核表.xlsx";

    static String projectDataFilePath = "6个QA文档/替换数据汇总表.xlsx";
    static String sheetName2 = "研发审计项目总表数据";

    public static void main(String[] args) throws Exception{
        Data data = getData(filePath, sheetName, projectDataFilePath, sheetName2);
        writeExcel(data);
    }

    private static void writeExcel(Data data) throws Exception{
        Set<String> set = new HashSet<>();

        Map<String, Project> projectMap = data.getProjectMap();
        for (String key : projectMap.keySet()) {
            if(!key.equals("智慧停车巡检APP")) continue;
            Project project = projectMap.get(key);
            System.out.println(JSON.toJSON(project));
            String format = String.format(targetFilePath, key);
            set.add(format);
            FileInputStream inputStream = new FileInputStream(templatePath);
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.cloneSheet(0, data.yuefen);


            Cell cell = sheet.getRow(0).getCell(0);
            cell.setCellValue(String.format(a1, data.year, data.yuefen));
            Cell cell1 = sheet.getRow(1).getCell(0);
            cell1.setCellValue(String.format(a2,project.projectName, project.负责人, project.projectCode));

            CellStyle cellType = cell1.getCellStyle();

            List<Person> personList = project.getPersonList();
            sheet.shiftRows(4, 4+personList.size(), personList.size(), true, false);

            int num = 0;
            for (int i = 0; i < personList.size(); i++) {
                Person person = personList.get(i);
                XSSFRow row = sheet.createRow(i + 4);
                XSSFCell cell2 = row.createCell(0);
                cell2.setCellStyle(cellType);
                cell2.setCellValue(person.role);
                XSSFCell cell3 = row.createCell(1);
                cell3.setCellStyle(cellType);
                cell3.setCellValue(person.name);

                num = 0;
                for (Map.Entry<String, String> entry : person.data.entrySet()) {
                    String key1 = entry.getKey();
                    String value = entry.getValue();
                    SimpleDateFormat sf = new SimpleDateFormat("yyyy/MM/dd");
                    sheet.getRow(2).getCell(num+2).setCellValue(sf.parse(key1));
                    XSSFCell cell4 = row.createCell(num + 2);
                    cell4.setCellStyle(cellType);
                    cell4.setCellValue(Double.parseDouble(value));
                    num++;
                }
            }
            workbook.write(new FileOutputStream(new File(format)));
            workbook.close();

            Workbook workbook1 = new Workbook();
            workbook1.loadFromFile(format);
            Worksheet worksheet = workbook1.getWorksheets().get(data.yuefen);
            System.out.println(num);
            worksheet.deleteColumn(num+2+0);
            worksheet.deleteColumn(num+2+1);
            worksheet.deleteColumn(num+2+2);
        }
        for (String s : set) {
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(new File(s)));
            workbook.removeSheetAt(0);
            workbook.write(new FileOutputStream(new File(s)));
            workbook.close();
        }
    }


    private static Data getData(String filePath, String  sheetName,
                                String projectDataFilePath,String sheetName2) throws IOException {
        File file = new File(filePath);

        Data data = new Data();
        String name = file.getName();
        String year = name.substring(0, 4);
        data.year = year;
        data.yuefen = Integer.parseInt(name.substring(4,6)) + "月";

        Map<String, Map<String, String>> projectData = getProjectData(projectDataFilePath, sheetName2);

        FileInputStream inputStream = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheet(sheetName);

        for (int i = 1; i < sheet.getLastRowNum(); i++) {
            XSSFRow row = sheet.getRow(i);

            String projectName = row.getCell(2).getStringCellValue();
            if(StringUtils.isNotBlank(projectName)){
                Project project = data.projectMap.getOrDefault(projectName, new Project());
                project.projectName = projectName;
                if(projectData.containsKey(projectName)){
                    project.projectCode = projectData.get(projectName).get("项目编号");
                    project.负责人 = projectData.get(projectName).get("项目负责人");
                }
                Person person = new Person();
                person.name = row.getCell(1).getStringCellValue();
                person.role = row.getCell(0).getStringCellValue();

                XSSFRow row0 = sheet.getRow(0);
                for (int j = 3; j < row0.getLastCellNum(); j++) {
                    String key = row0.getCell(j).getStringCellValue();
                    String value = row.getCell(j)==null? "0" : "" + (int)row.getCell(j).getNumericCellValue();
                    person.data.put(key, value);
                }
                project.personList.add(person);
                data.projectMap.put(projectName, project);
            }
        }
        return data;
    }

    private static Map<String, Map<String,String>> getProjectData(String filePath, String sheetName) throws IOException {
        XSSFWorkbook projectData = new XSSFWorkbook(new FileInputStream(new File(filePath)));
        XSSFSheet sheet = projectData.getSheet(sheetName);
        Map<String, Map<String,String>> map = new HashMap<>();
        for (int i = 1; i < sheet.getLastRowNum(); i++) {
            XSSFRow row = sheet.getRow(i);
            XSSFRow header = sheet.getRow(0);
            HashMap<String,String> rowData = new HashMap<>();
            map.put(row.getCell(0).getStringCellValue(), rowData);
            for (int j = 0; j < row.getLastCellNum(); j++) {
                rowData.put(header.getCell(j).getStringCellValue(), ExcelUtil.getValue(row.getCell(j)));
            }
        }
        return map;
    }


    @lombok.Data
    public static class Data{

        String yuefen;
        String year;
        Map<String,Project> projectMap = new HashMap<>();
    }

    @lombok.Data
    static class Project{
        String projectName;
        String projectCode = "";
        String 负责人 = "";
        List<Person> personList = new ArrayList<>();
    }
    @lombok.Data
    static class Person{
        String name;
        String role;
        LinkedHashMap<String,String> data = new LinkedHashMap<>();

    }


}
