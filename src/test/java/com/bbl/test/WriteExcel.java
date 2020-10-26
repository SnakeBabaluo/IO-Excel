package com.bbl.test;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 创建日期 : 2020/10/26 20:32
 *
 * 写出excel文件
 *
 * 文件格式
 * XSSFWorkbook：工作簿
 * XSSFSheet：工作表
 * XSSFRow：行
 * XSSFCell：单元格
 */
public class WriteExcel {

    @Test
    public void createExcel() throws IOException{
        //创建工作簿
        Workbook wk = new XSSFWorkbook();
        //创建工作表
        Sheet sheetAt = wk.createSheet("测试写Excel");
        //创建行,默认0行开始
        Row row = sheetAt.createRow(0);
        //使用行创建单元格,也是默认从0开始
        Cell cell = row.createCell(0);
        //给单元格赋值
        //第1行的第1个单元格放入 "姓名"
        cell.setCellValue("姓名");
        //第1行的第2个格写入 "年龄"
        row.createCell(1).setCellValue("年龄");
        row.createCell(2).setCellValue("所在地");

        //切换第二行
        row=sheetAt.createRow(1);
        //第2行的第1个单元格放入 "小明"
        row.createCell(0).setCellValue("小明");
        row.createCell(1).setCellValue(20);
        row.createCell(2).setCellValue("北京");

        //切换第3行
        row=sheetAt.createRow(2);
        //第3行的第1个单元格放入 "小李"
        row.createCell(0).setCellValue("小李");
        row.createCell(1).setCellValue(30);
        row.createCell(2).setCellValue("南京");
        //写出文件
        wk.write(new FileOutputStream(new File("d:\\createExcel.xlsx")));

        //关闭工作簿
        wk.close();

    }
}
