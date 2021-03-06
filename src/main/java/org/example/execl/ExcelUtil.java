package org.example.execl;

import lombok.Data;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import static org.example.execl.派工单App.*;

public class ExcelUtil {

    public static void writeExcel(String targetFilePath, Project project) throws IOException {
        System.out.println(project.name);
        FileInputStream inputStream = new FileInputStream(templatePath);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

        for (Map.Entry<String, List<YueShuJu>> map : project.yueShuJu.entrySet()) {
            System.out.println(map.getKey());
            // 复制模版
            XSSFSheet sheet = workbook.cloneSheet(1, map.getKey());
            // 修改a1
            Cell cell = sheet.getRow(0).getCell(0);
            cell.setCellValue(String.format(a1, map.getKey()));
            //
            Cell cell1 = sheet.getRow(1).getCell(0);
            cell1.setCellValue(String.format(a2,project.name, project.projectCode));

            List<YueShuJu> value = map.getValue();
            int size = value.size();
            sheet.shiftRows(3, 3+size, size);

            for (int i = 0; i < size; i++) {
                YueShuJu yueShuJu = map.getValue().get(i);
                XSSFRow row = sheet.createRow(i+3);
                row.createCell(0).setCellValue(yueShuJu.index);
                row.createCell(1).setCellValue(yueShuJu.name);
                row.createCell(2).setCellValue(yueShuJu.sex);
                row.createCell(3).setCellValue(yueShuJu.部门);
                row.createCell(4).setCellValue(yueShuJu.项目角色);
            }
        }
        workbook.removeSheetAt(1);
        workbook.write(new FileOutputStream(
                new File(String.format(targetFilePath, project.name))));
    }

    // 读取EXCEL数据
    static List<Project> readExcel(String filePath, String sheetName, boolean isFilter) throws Exception {
        FileInputStream inputStream = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(inputStream);

        Map<String,String> map = getProjectCodeMap(workbook.getSheet(sheetName2));

        Sheet sheet = workbook.getSheet(sheetName);
        int lastRowNum = sheet.getLastRowNum();
        List<Project> projects = new ArrayList<>();

        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if(region.getLastRow() == 0){
                String value = sheet.getRow(region.getFirstRow()).getCell(region.getFirstColumn()).getStringCellValue();

                Project project = new Project();
                projects.add(project);
                project.name = value;
                project.projectCode = map.get(value);
                for (int j = region.getFirstColumn(); j <= region.getLastColumn(); j++) {
                    String yuefen = sheet.getRow(region.getFirstRow()+1).getCell(j).getStringCellValue();
                    List<YueShuJu> yueShuJus = new ArrayList<>();
                    int index = 1;
                    for (int k = region.getFirstRow() + 2; k < lastRowNum; k++) {
                        double number = sheet.getRow(k).getCell(j).getNumericCellValue();
                        if(number < 0 & isFilter){
                            continue;
                        }
                        String 部门 = getValue(sheet.getRow(k).getCell(0));    // 部门
                        String 项目角色 = getValue(sheet.getRow(k).getCell(3)); // 项目角色
                        String 性别 = getValue(sheet.getRow(k).getCell(4)); // 性别
                        String 姓名 = getValue(sheet.getRow(k).getCell(5)); // 姓名
                        YueShuJu yueShuJu = new YueShuJu(index, 姓名, 性别, 部门, 项目角色, number);
                        yueShuJus.add(yueShuJu);
                        index++;
                    }
                    if(yueShuJus.size() > 0){
                        project.yueShuJu.put(yuefen, yueShuJus);
                    }
                }
            }
        }
        return projects;
    }

    private static Map<String, String> getProjectCodeMap(Sheet sheet) {
        Map<String, String> result = new HashMap<>();
        for (int i = 1; i < sheet.getLastRowNum(); i++) {
            String name = sheet.getRow(i).getCell(0).getStringCellValue();
            if(StringUtils.isNotBlank(name)){
                String code = sheet.getRow(i).getCell(2).getStringCellValue();
                result.put(name, code);
            }
        }
        return result;
    }

    public static Map<String, Map<String,String>> getProjectData(String filePath, String sheetName) throws IOException {
        XSSFWorkbook projectData = new XSSFWorkbook(new FileInputStream(new File(filePath)));
        XSSFSheet sheet = projectData.getSheet(sheetName);
        Map<String, Map<String,String>> map = new HashMap<>();
        for (int i = 1; i < sheet.getLastRowNum(); i++) {
            XSSFRow row = sheet.getRow(i);
            XSSFRow header = sheet.getRow(0);
            HashMap<String,String> rowData = new HashMap<>();
            map.put(row.getCell(0).getStringCellValue(), rowData);
            for (int j = 0; j < row.getLastCellNum(); j++) {
                String key = header.getCell(j).getStringCellValue();
                XSSFCell cell = row.getCell(j);
                if(key.equals("立项时间") ||
                        key.equals("进入开发时间") ||
                        key.equals("完成开发时间")){
                    Date dateCellValue = cell.getDateCellValue();
                    String val = "";
                    if(dateCellValue != null) {
                        SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
                        val =  sf.format(dateCellValue);
                    }
                    rowData.put(key, val);
                }else{
                    rowData.put(key, ExcelUtil.getValue(cell));
                }
            }
        }
        return map;
    }

    public static String getValue(Cell cell) {
        String val = "";
        switch (cell.getCellType()) {
            case STRING:   // 字符串类型
                val = cell.getStringCellValue().trim();
                break;
            case NUMERIC:  // 数值类型
                val = new DecimalFormat("#").format(cell.getNumericCellValue());
                break;
            default: //其它类型
                break;
        }
        return val;
    }

    public static void removeColumn(Sheet sheet, int column) {
        for (Row row : sheet) {
            Cell cell = row.getCell(column);
            if (cell != null) {
                row.removeCell(cell);
            }
        }
    }


    @Data
    static class Project{
        String name;
        String projectCode;
        LinkedHashMap<String,List<YueShuJu>> yueShuJu = new LinkedHashMap<>();
    }

    @Data
    static class YueShuJu{
        int index;
        String name;
        String sex;
        String 部门;
        String 项目角色;
        double number;

        public YueShuJu(int index, String name, String sex, String 部门, String 项目角色, double number) {
            this.index = index;
            this.name = name;
            this.sex = sex;
            this.部门 = 部门;
            this.项目角色 = 项目角色;
            this.number = number;
        }

        @Override
        public String toString() {
            return "YueShuJu{" +
                    "index=" + index +
                    ", name='" + name + '\'' +
                    ", sex='" + sex + '\'' +
                    ", 部门='" + 部门 + '\'' +
                    ", 项目角色='" + 项目角色 + '\'' +
                    '}';
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;

            YueShuJu yueShuJu = (YueShuJu) o;

            if (name != null ? !name.equals(yueShuJu.name) : yueShuJu.name != null) return false;
            if (sex != null ? !sex.equals(yueShuJu.sex) : yueShuJu.sex != null) return false;
            if (部门 != null ? !部门.equals(yueShuJu.部门) : yueShuJu.部门 != null) return false;
            return 项目角色 != null ? 项目角色.equals(yueShuJu.项目角色) : yueShuJu.项目角色 == null;
        }

        @Override
        public int hashCode() {
            int result = name != null ? name.hashCode() : 0;
            result = 31 * result + (sex != null ? sex.hashCode() : 0);
            result = 31 * result + (部门 != null ? 部门.hashCode() : 0);
            result = 31 * result + (项目角色 != null ? 项目角色.hashCode() : 0);
            return result;
        }
    }
}
