package com.amosannn.core;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

/**
 * @author amos.lin
 *
 */
public class TXTParser {


    /**
     * 将Excel文件转换为TXT文件
     * @param filePath 输入的Excel文件路径
     * @param outputPath 输出的TXT文件路径
     * @throws Exception
     */
    @Test
    public static void toTXT(final String filePath, final String outputPath) throws Exception{
        // 文件名与文件路径预处理
        final int fileNameIndex = filePath.lastIndexOf("\\");
        final String fileTypeSub = filePath.substring(fileNameIndex+1);
        final String[] fileType = fileTypeSub.split("\\.");
        final InputStream is = new FileInputStream(filePath);
        final StringBuilder strBuilder = new StringBuilder();


        // 识别xls与xlsx后缀名
        if(fileType[1].equals("xls")){
            final HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
            HSSFSheet sheet = null;
            HSSFRow row = null;
            HSSFCell cell = null;
            sheet = hssfWorkbook.getSheetAt(0);
            row = sheet.getRow(0);
            cell = row.getCell(0);
            new TXTParser().xlsParser(hssfWorkbook, sheet, row, cell, strBuilder, outputPath);
        } if (fileType[1].equals("xlsx")){
            XSSFWorkbook xssfWorkbook = xssfWorkbook = new XSSFWorkbook(is);
            XSSFSheet sheet = null;
            XSSFRow row = null;
            XSSFCell cell = null;
            sheet = xssfWorkbook.getSheetAt(0);
            row = sheet.getRow(0);
            cell = row.getCell(0);
            new TXTParser().xlsxParser(xssfWorkbook, sheet, row, cell, strBuilder, outputPath);
        }

    }

    /**
     * xls格式转换器
     * @param wb
     * @param sheet
     * @param row
     * @param cell
     * @param strBuilder
     * @param outputPath
     * @throws IOException
     */
    public void xlsParser(final HSSFWorkbook wb, HSSFSheet sheet, HSSFRow row, HSSFCell cell, final StringBuilder strBuilder, final String outputPath) throws IOException{
        String result="";
        //三层循环遍历Excel表格
        //遍历sheet
        for(int sheetNum = 0; sheetNum < wb.getNumberOfSheets(); sheetNum++){
            sheet = wb.getSheetAt(sheetNum);
            if(null == sheet) {
                continue;
            }

            strBuilder.append(wb.getSheetName(sheetNum)+"/t");
            //遍历row
            // getLastRowNum() 最后一行行标，比行数小1
            for(int rowNum = 0; rowNum < sheet.getLastRowNum()+1; rowNum++){
                row = sheet.getRow(rowNum);

                if(null == row) {
                    continue;
                }

                //遍历cell
                // getLastCellNum() 获取列数，比最后一列列标大1
                for(int colNum = 0; colNum < row.getLastCellNum(); colNum++){
                    cell = row.getCell(colNum);
                    if(null == cell||cell.getCellType()==cell.CELL_TYPE_BLANK) {
                        continue;
                    }
                    strBuilder.append(getCellText(cell)+"\t");

//                    final String temp = getCellText(cell);
//                    System.out.print(temp + " ");
                    strBuilder.append("\n");
                }
            }
            result+=strBuilder.toString();
        }
        byte[] bytes=new byte[]{};
        bytes=result.getBytes();
        final OutputStream os = new FileOutputStream(new File(outputPath));
        os.write(bytes);
        os.close();
        wb.close();
    }

    /**
     * xlsx格式转换器
     * @param wb
     * @param sheet
     * @param row
     * @param cell
     * @param strBuilder
     * @param outputPath
     * @throws IOException
     */
    public void xlsxParser(final XSSFWorkbook wb, XSSFSheet sheet, XSSFRow row, XSSFCell cell, final StringBuilder strBuilder, final String outputPath) throws IOException{
        String result="";
        //三层循环遍历Excel表格
        //遍历sheet
        for(int sheetNum = 0; sheetNum < wb.getNumberOfSheets(); sheetNum++){
            sheet = wb.getSheetAt(sheetNum);
            if(null == sheet) {
                continue;
            }

            strBuilder.append(wb.getSheetName(sheetNum)+"/t");
            //遍历row
            // getLastRowNum() 最后一行行标，比行数小1
            for(int rowNum = 0; rowNum < sheet.getLastRowNum()+1; rowNum++){
                row = sheet.getRow(rowNum);

                if(null == row) {
                    continue;
                }

                //便利cell
                // getLastCellNum() 获取列数，比最后一列列标大1
                for(int colNum = 0; colNum < row.getLastCellNum(); colNum++){
                    cell = row.getCell(colNum);
                    if(null == cell||cell.getCellType()==cell.CELL_TYPE_BLANK) {
                        continue;
                    }
                    strBuilder.append(getCellText(cell)+"\t");

//                    final String temp = getCellText(cell);
//                    System.out.print(temp + " ");
                    strBuilder.append("\n");
                }
            }
            result+=strBuilder.toString();
        }
        byte[] bytes=new byte[]{};
        bytes=result.getBytes();
        final OutputStream os = new FileOutputStream(new File(outputPath));
        os.write(bytes);
        os.close();
        wb.close();
    }

    /**
     * 表格内容获取器
     * @param cell
     * @return
     */
    @SuppressWarnings("deprecation")
    public static String getCellText(final Cell cell){
        String cellText = null;
        switch (cell.getCellType()) {
        case Cell.CELL_TYPE_BOOLEAN:
            cellText = cell.getBooleanCellValue()+"";
            break;
        case Cell.CELL_TYPE_FORMULA:
            cellText = cell.getDateCellValue()+"";
            break;
        case Cell.CELL_TYPE_NUMERIC:
            cellText = cell.getNumericCellValue()+"";
            break;
        case Cell.CELL_TYPE_STRING:
            cellText = cell.getStringCellValue()+"";
            break;
        default:
            break;
        }
        return  cellText;
    }

}
