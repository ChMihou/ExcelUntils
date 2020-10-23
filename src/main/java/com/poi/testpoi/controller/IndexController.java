package com.poi.testpoi.controller;

import com.poi.testpoi.pojo.User;
import com.poi.testpoi.service.UserService;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

@Controller
public class IndexController {

    @Autowired
    private UserService userService;


    @RequestMapping("/index")
    public String showUser(Model model) {
        List<User> users = userService.selectUsers();
        model.addAttribute("user", users);
        return "index";
    }

    /**
     * （HSSFWORKbook是2003版本Excel用的工具类
     * 不同版本的Excel工具类也不同）
     *
     * @param response
     * @throws IOException
     */
    @RequestMapping(value = "/export")
    @ResponseBody
    public void export(HttpServletResponse response) throws IOException {
        List<User> users = userService.selectUsers();

        HSSFWorkbook wb = new HSSFWorkbook();

        //编辑excel表格的页脚
        HSSFSheet sheet = wb.createSheet("获取excel测试表格");
        //创建第一个单元格
        HSSFRow row = sheet.createRow(0);
        //设置行高
        row.setHeight((short) (26.25 * 20));
        //为第一行单元格设值
        row.createCell(0).setCellValue("用户信息列表");

        /*为标题设计空间
         * firstRow从第1行开始
         * lastRow从第1行结束
         *从第1个单元格开始
         *从第3个单元格结束
         */
        CellRangeAddress rowRegion = new CellRangeAddress(0, 0, 0, 2);
        sheet.addMergedRegion(rowRegion);

		/*CellRangeAddress columnRegion = new CellRangeAddress(1,4,0,0);
		sheet.addMergedRegion(columnRegion);*/


        /*
         * 动态获取数据库列 sql语句 select COLUMN_NAME from INFORMATION_SCHEMA.Columns where table_name='user' and table_schema='test'
         * 第一个table_name 表名字
         * 第二个table_name 数据库名称
         * */
        //创建第二行表格
        row = sheet.createRow(1);
        //设置行高
        row.setHeight((short) (22.50 * 20));
        //为第一个单元格设值
        row.createCell(0).setCellValue("用户Id");
        //为第二个单元格设值
        row.createCell(1).setCellValue("用户名");
        //为第三个单元格设值
        row.createCell(2).setCellValue("用户密码");

        for (int i = 0; i < users.size(); i++) {
            //创建第i+2行单元表格
            row = sheet.createRow(i + 2);
            User user = users.get(i);
            row.createCell(0).setCellValue(user.getUid());
            row.createCell(1).setCellValue(user.getUsername());
            row.createCell(2).setCellValue(user.getPassword());
        }
        //设置默认行高
        sheet.setDefaultRowHeight((short) (16.5 * 20));
        //列宽自适应，
        for (int i = 0; i <= users.size() + 2; i++) {
            sheet.autoSizeColumn(i);
        }
        /**
         * 下载到当前页面
         */
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        OutputStream os = response.getOutputStream();
        //默认Excel名称
        response.setHeader("Content-disposition", "attachment;filename=users.xls");
        wb.write(os);
        /**
         * 下载到指定文件夹内
         */
        String resultName = "";
        String ctxPath = "D://upFiles";
        String name = new SimpleDateFormat("ddHHmmss").format(new Date());
        String fileName = name + "users.xlsx";
        String bizPath = "files";
        String nowday = new SimpleDateFormat("yyyyMMdd").format(new Date());
        File file = new File(ctxPath + File.separator + bizPath + File.separator + nowday);
        if (!file.exists()) {
            file.mkdirs();// 创建文件根目录
        }
        String savePath = file.getPath() + File.separator + fileName;
        resultName = bizPath + File.separator + nowday + File.separator + fileName;
        if (resultName.contains("\\")) {
            resultName = resultName.replace("\\", "/");
        }
        System.out.print(resultName);
        System.out.print(savePath);
        // 响应到客户端需要下面注释的代码
//            this.setResponseHeader(response, filename);
//            OutputStream os = response.getOutputStream(); //响应到服务器
        // 保存到当前路径savePath
        os = new FileOutputStream(savePath);

        wb.write(os);
        os.flush();
        os.close();
    }


    @RequestMapping(value = "/import")
    public String exImport(@RequestParam(value = "filename") MultipartFile file, HttpSession session) {

        boolean a = false;

        String fileName = file.getOriginalFilename();

        try {
            a = userService.batchImport(fileName, file);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "redirect:index";
    }


    /**
     * 获取样式
     *
     * @param hssfWorkbook
     * @param styleNum
     * @return
     */
    public HSSFCellStyle getStyle(HSSFWorkbook hssfWorkbook, Integer styleNum) {
        HSSFCellStyle style = hssfWorkbook.createCellStyle();
        style.setBorderRight(BorderStyle.THIN);//右边框
        style.setBorderBottom(BorderStyle.THIN);//下边框

        HSSFFont font = hssfWorkbook.createFont();
        font.setFontName("微软雅黑");//设置字体为微软雅黑

        HSSFPalette palette = hssfWorkbook.getCustomPalette();//拿到palette颜色板,可以根据需要设置颜色
        switch (styleNum) {
            case (0): {//HorizontalAlignment
                style.setAlignment(HorizontalAlignment.CENTER_SELECTION);//跨列居中
                font.setBold(true);//粗体
                font.setFontHeightInPoints((short) 14);//字体大小
                style.setFont(font);
                palette.setColorAtIndex(HSSFColor.BLUE.index, (byte) 184, (byte) 204, (byte) 228);//替换颜色板中的颜色
                style.setFillForegroundColor(HSSFColor.BLUE.index);
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            break;
            case (1): {
                font.setBold(true);//粗体
                font.setFontHeightInPoints((short) 11);//字体大小
                style.setFont(font);
            }
            break;
            case (2): {
                font.setFontHeightInPoints((short) 10);
                style.setFont(font);
            }
            break;
            case (3): {
                style.setFont(font);

                palette.setColorAtIndex(HSSFColor.GREEN.index, (byte) 0, (byte) 32, (byte) 96);//替换颜色板中的颜色
                style.setFillForegroundColor(HSSFColor.GREEN.index);
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            break;
        }

        return style;
    }


}
