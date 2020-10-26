package com.bbl.test;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.IOException;

/**
 * 创建日期 : 2020/10/26 20:02
 */
public class ReadForExcel {

    @Test
    public void  readForExcel() throws IOException {
        //创造工作簿,构造方法文件路径
        XSSFWorkbook wk = new XSSFWorkbook("d:\\aa.xlsx");
        //获取工作表,下标是从0开始
        Sheet sht = wk.getSheetAt(0);
        //获所有的行,下标也是从0开始的
        for (Row cells : sht) {
            //遍历当前行
            for (Cell cell : cells) {
                int cellType = cell.getCellType();
                //这个常量代表着数值类型, 详情点进去看
                if (Cell.CELL_TYPE_NUMERIC == cellType) {
                    System.out.print(cell.getNumericCellValue() + ",");
                }else {
                    System.out.print(cell.getStringCellValue()+",");
                }
            }
            System.out.println();
        }
        //关闭工作簿
        wk.close();
    }
}
