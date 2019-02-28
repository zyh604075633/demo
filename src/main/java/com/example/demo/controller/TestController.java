package com.example.demo.controller;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.net.URLEncoder;

@Controller
@RequestMapping("/test")
public class TestController {
    @RequestMapping("/test")
    @ResponseBody
    public String test(HttpServletResponse response, HttpServletRequest request) throws Exception {
        String contextPath = request.getServletContext().getContextPath();
        System.out.println(contextPath);
        return "测试类";
    }
    @RequestMapping("/download")
    public void download(HttpServletResponse response, HttpServletRequest request) throws Exception {
        SXSSFWorkbook sworkbook = new SXSSFWorkbook();
        Sheet sheet = sworkbook.createSheet("工作簿名称");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("这是0.0单元格");
        Cell cell1 = row.createCell(1);
        cell1.setCellValue("这是0.1单元格");//单元格赋值


        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet createSheet = workbook.createSheet("工作簿名称");
        XSSFRow createRow = createSheet.createRow(0);
        XSSFCell createCell = createRow.createCell(0);
        createCell.setCellValue("这是0.0单元格");
        XSSFCell createCell1 = createRow.createCell(1);
        createCell1.setCellValue("这是0.1单元格");//单元格赋值

        //设置content-disposition响应头控制浏览器以下载的形式打开文件，中文文件名要使用URLEncoder.encode方法进行编码，否则会出现文件名乱码
        response.setHeader("content-disposition", "attachment;filename="+ URLEncoder.encode("这是热乎的下载文件.xlsx", "UTF-8"));

        OutputStream out = response.getOutputStream();
        sworkbook.write(out);
        out.flush();
        out.close();
    }
}
