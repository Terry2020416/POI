package cn.blb.test;

import cn.blb.entity.Animal;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class POICellStyleDemo {
    @Test
    public void createExcel() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("哈哈");

        XSSFRow row = sheet.createRow(0);//第1行
        row.createCell(0).setCellValue("id");
        row.createCell(1).setCellValue("name");
        row.createCell(2).setCellValue("kind");
        row.createCell(3).setCellValue("age");

        List<Animal> list = new ArrayList<Animal>();
        list.add(new Animal(1,"xx1","0",3));
        list.add(new Animal(2,"xx2","1",2));
        list.add(new Animal(3,"xx3","0",3));

        for(int i =0;i<list.size();i++){
            XSSFRow row1 = sheet.createRow(i+1);
            row1.createCell(0).setCellValue(list.get(i).getId());
            row1.createCell(1).setCellValue(list.get(i).getName());
            row1.createCell(2).setCellValue(list.get(i).getKind());
            row1.createCell(3).setCellValue(list.get(i).getAge());

        }
//        XSSFCell cell = row.createCell(0);
//        cell.setCellValue("编号");
//        row.createCell(1).setCellValue("姓名");
//        row.createCell(2).setCellValue("年龄");
//        row.createCell(3).setCellValue("性别");
//        Font font = workbook.createFont();
//        font.setItalic(true);
//        font.setFontName("微软雅黑");
//
//        CellStyle cellStyle = workbook.createCellStyle();
//        cellStyle.setFont(font);
//        cellStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
//        cellStyle.setFillPattern(CellStyle.BIG_SPOTS);



        FileOutputStream fos = new FileOutputStream("C:\\Users\\Administrator.KQ3XM23MUKVL0OJ\\Desktop\\Test.xlsx");

        workbook.write(fos);
        fos.close();
        workbook.close();
    }





    @Test
    public void setCellFontStyle() throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("单元格样式设置表");
        XSSFRow row = sheet.createRow(1);//获取第二行

        //设置单元格样式
        XSSFCell cell = row.createCell(0);//获取第二行第一个单元格
        cell.setCellValue("文本样式设置");

        //创建字体样式对象
        Font font = wb.createFont();
        font.setBold(true);//加粗
        font.setFontName("微软雅黑");//字体样式
        font.setItalic(true);//倾斜

        //设置样式
        CellStyle style = wb.createCellStyle();
        //将字体样式设置到单元格样式对象上
        style.setFont(font);

        //设置背景色
        style.setFillForegroundColor(IndexedColors.RED.getIndex());
        //style.setFillPattern(CellStyle.BIG_SPOTS);//填充方式：spots斑点
        style.setFillPattern(CellStyle.ALIGN_FILL);//填充方式：spots斑点


        cell.setCellStyle(style);

        FileOutputStream fos = new FileOutputStream("c:\\Users\\Administrator.KQ3XM23MUKVL0OJ\\Desktop\\POI设置样式Excle222.xlsx");
        wb.write(fos);
        fos.close();
        wb.close();

    }

    @Test
    public void setCellDataStyle() throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("单元格样式设置表");
        XSSFRow row = sheet.createRow(1);//获取第二行

        //设置单元格样式
        XSSFCell cell = row.createCell(0);//获取第二行第一个单元格
        XSSFCell cell2 = row.createCell(1);//获取第二行第一个单元格
        cell.setCellValue(100.1234);
        cell2.setCellValue(new Date());
        //设置单元格1的样式
        XSSFCellStyle cellStyle = wb.createCellStyle();
        XSSFDataFormat format = wb.createDataFormat();
        cellStyle.setDataFormat(format.getFormat("0.0"));

        //设置单元格2的样式--设置时间的显示
        XSSFCellStyle cellStyle2 = wb.createCellStyle();
        CreationHelper creationHelper = wb.getCreationHelper();
        cellStyle2.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd hh:mm:ss"));


        cell.setCellStyle(cellStyle);
        cell2.setCellStyle(cellStyle2);

        FileOutputStream fos = new FileOutputStream("c:\\Users\\Administrator.KQ3XM23MUKVL0OJ\\Desktop\\POI设置单元格文本样式Excle.xlsx");
        wb.write(fos);
        fos.close();
        wb.close();

    }


    /**
     * 设置单元格的文本样式
     * @throws Exception
     */
    @Test
    public void setCellFontStlyle() throws Exception {
        XSSFWorkbook wb= new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("单元格样式设置表");
        XSSFRow row = sheet.createRow(4);//获取第二行
        row.setHeightInPoints(50);//设置当前行高，便于查看对齐方式
        //设置单元格样式
        XSSFCell cell = row.createCell(4);//第二行的第五个单元格，设置他的样式
        cell.setCellValue("文本样式设置");
        //------------------------
        //创建字体样式对象
        Font fontSytle = wb.createFont();
        fontSytle.setFontName("黑体");//字体
        fontSytle.setItalic(true);//倾斜
        //创建单元格样式对象
        CellStyle style  = wb.createCellStyle();
        //1.将字体样式设置到单元格样式对象上
        style.setFont(fontSytle);
        //2.背景颜色
        //style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        //style.setFillPattern(CellStyle.BIG_SPOTS); //填充方式：spots斑点
        //3. 设置边框
        style.setBorderBottom(CellStyle.BORDER_DOUBLE);//底部边框样式
        style.setBottomBorderColor(IndexedColors.YELLOW.getIndex());//边框颜色
        style.setBorderLeft(CellStyle.BORDER_DASHED);//底部边框样式
        style.setLeftBorderColor(IndexedColors.RED.getIndex());//边框颜色
        //4. 对齐方式
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);//水平居中
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直居中*/
        //------------------------
        //给单元格设置样式
        cell.setCellStyle(style);
        FileOutputStream fos = new FileOutputStream("C:\\Users\\Administrator.KQ3XM23MUKVL0OJ\\Desktop\\文本样式设置.xlsx");
        wb.write(fos);
        wb.close();
        fos.close();
        System.out.println("设置完毕");
    }
}
