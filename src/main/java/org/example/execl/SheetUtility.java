package org.example.execl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class SheetUtility {


    public static void deleteColumn( Sheet sheet, int columnToDelete ){
        for (int rId = 0; rId <= sheet.getLastRowNum(); rId++) {
            Row row = sheet.getRow(rId);
            if(row!=null){
                Cell cell = row.getCell(columnToDelete);
                if(cell!=null){
                    row.removeCell(cell);
                }
            }
        }
    }

    private static void cloneCell( Cell cNew, Cell cOld ){
        cNew.setCellComment( cOld.getCellComment() );
        cNew.setCellStyle( cOld.getCellStyle() );

        switch ( cNew.getCellType() ){
            case BOOLEAN:{
                cNew.setCellValue( cOld.getBooleanCellValue() );
                break;
            }
            case NUMERIC:{
                cNew.setCellValue( cOld.getNumericCellValue() );
                break;
            }
            case STRING:{
                cNew.setCellValue( cOld.getStringCellValue() );
                break;
            }
            case ERROR:{
                cNew.setCellValue( cOld.getErrorCellValue() );
                break;
            }
            case FORMULA:{
                cNew.setCellFormula( cOld.getCellFormula() );
                break;
            }
        }

    }
}