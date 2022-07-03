package org.example.doc;

import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class DocUtil {

    public static final String mark = "[%s]";

    // 替换文本
    static void replaceDoc(String template, String format, List<List<Pair<String, String>>> projectData) throws Exception {
        for (List<Pair<String,String>> data : projectData) {
            String projectName = data.get(0).getValue();
            String out = String.format(format, projectName);
            Function function = (xml) -> {
                for (Pair<String, String> datum : data) {
                    if(xml.contains(datum.getKey())){
                        System.out.println(datum.getKey() + "-> " + datum.getValue());
                        xml = xml.replace(datum.getLeft(), datum.getRight());
                    }

                }
                return xml;
            };
            editDocx(template, out, function);
        }
    }

    // 读取EXCEL数据
    static  List<List<Pair<String,String>>> readExcel(String filePath, String sheetName) throws Exception {
        FileInputStream inputStream = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheet(sheetName);
        List<List<Pair<String,String>>>  result = new ArrayList<>();
        for (int i = 1; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            Row keyRow = sheet.getRow(0);
            List<Pair<String,String>> rowKeyWord= new ArrayList<>();
            for (int j = 0; j < row.getLastCellNum(); j++) {
                String key = keyRow.getCell(j).getStringCellValue();
                String value = getValue(row.getCell(j), key.contains("时间"));
                rowKeyWord.add(Pair.of(getKey(key), value));
            }
            result.add(rowKeyWord);
        }
        return result;
    }

    private static String getValue(Cell cell, boolean  isDate) {
        String val = "";
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case STRING:   // 字符串类型
                val = cell.getStringCellValue().trim();
                break;
            case NUMERIC:  // 数值类型
                if(isDate){
                    SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy年MM月dd日");
                    val = simpleDateFormat.format(cell.getDateCellValue());
                }else{
                    val = new DecimalFormat("#").format(cell.getNumericCellValue());
                }
                break;
            default: //其它类型
                break;
        }
        return val;
    }


    // 编辑 docx
    public static void editDocx(String template,
                                String out,
                                Function function) throws Exception {
        FileInputStream templateFile = new FileInputStream(template);
        FileOutputStream outFile = new FileOutputStream(out);

        try (ZipInputStream zin = new ZipInputStream(templateFile);
             ZipOutputStream zos = new ZipOutputStream(outFile)) {
            ZipEntry entry;
            while ((entry = zin.getNextEntry()) != null) {
                //把输入流的文件传到输出流中 如果是word/document.xml由我们输入
                zos.putNextEntry(new ZipEntry(entry.getName()));
                if("word/document.xml".equals(entry.getName()) || "word/header1.xml".equals(entry.getName())){
                    System.out.println(entry.getName());
                    String xml = new BufferedReader(new InputStreamReader(zin))
                            .lines()
                            .collect(Collectors.joining(System.lineSeparator()));
                    System.out.println(xml);
                    xml = function.process(xml);

                    ByteArrayInputStream byteIn = new ByteArrayInputStream(xml.getBytes());
                    int c;
                    while ((c = byteIn.read()) != -1) {
                        zos.write(c);
                    }
                    byteIn.close();
                }else {
                    int c;
                    while ((c = zin.read()) != -1) {
                        zos.write(c);
                    }
                }
            }
        }
    }

    public static String getKey(String key){
        return String.format(mark, key);
    }

    @FunctionalInterface
    interface Function {
        String process(String xml);
    }
}
