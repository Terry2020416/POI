package cn.blb.test;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.*;

public class POIDemo2 {

    @Test
    public void readExcelDatas() throws IOException {
        //1、创建一个工作簿
        XSSFWorkbook workbook = new XSSFWorkbook("c:\\Users\\Administrator.KQ3XM23MUKVL0OJ\\Desktop\\TestPOI.xlsx");
        //2、获取工作表
        XSSFSheet sheet = workbook.getSheetAt(0);
        //3、获取行对象
        int lastRowNum = sheet.getLastRowNum();
        for(int rowIndex = 0; rowIndex <= lastRowNum;rowIndex++){
            //操作当前行对象
            XSSFRow row = sheet.getRow(rowIndex);
            if(row != null){
                //获取总列数据
                int lastColumnNum = row.getLastCellNum();
                for(int columnNum = 0; columnNum <= lastColumnNum; columnNum++){
                    //4、获取当前单元格(列)对象
                    XSSFCell cell = row.getCell(columnNum);
                    if(cell != null)
                        //5、获取数据
                        System.out.print(cell + " ");
                }
                System.out.println();
            }
        }

    }

    @Test
    public void createExcelSheets() throws IOException {
        //1、创建一个工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        //获取输出流对象
        FileOutputStream fileOutputStream = new FileOutputStream("c:\\Users\\Administrator.KQ3XM23MUKVL0OJ\\Desktop\\POI创建的Excle.xlsx");

        //2、创建工作表
        workbook.createSheet("宠物管理");
        workbook.createSheet("用户管理");
        workbook.createSheet("商品管理");

        //将工作簿输出到执行文件中，创建一个空Excel文件
        workbook.write(fileOutputStream);
    }

    @Test
    public void createExcelDatas() throws IOException {
        //1、创建一个工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        //获取输出流对象
        FileOutputStream fileOutputStream = new FileOutputStream("c:\\Users\\Administrator.KQ3XM23MUKVL0OJ\\Desktop\\POI创建的Excle.xlsx");

        //2、创建工作表
        workbook.createSheet("宠物管理");
        workbook.createSheet("用户管理");
        XSSFSheet sheet = workbook.createSheet("商品管理");

        //3、创建行(标题)
        XSSFRow row = sheet.createRow(0);

        //4、创建列
        row.createCell(0).setCellValue("id");
        row.createCell(1).setCellValue("name");
        row.createCell(2).setCellValue("age");
        row.createCell(3).setCellValue("isAdmin");

        //创建行(内容)
        for (int i = 0; i <= 100; i++){
            XSSFRow row1 = sheet.createRow(i);
            row1.createCell(0).setCellValue(1000+i);
            row1.createCell(1).setCellValue("Lucky"+i);//设置数值类型
            row1.createCell(2).setCellValue(0+i);//设置整行
            row1.createCell(3).setCellValue(i%2==0?true:false);//设置布尔类型
        }
        //将工作簿输出到执行文件中，创建一个空Excel文件
        workbook.write(fileOutputStream);

        workbook.close();
        fileOutputStream.close();
    }

    /**
     * 一次性读取Excle数据
     * @throws IOException
     */
    @Test
    public void readExcelDatasOnce() throws IOException {
        //1、获取Excle输入流对象
        InputStream is = new FileInputStream("c:\\Users\\Administrator.KQ3XM23MUKVL0OJ\\Desktop\\POI创建的Excle.xlsx");
        //2、POI提供了操作流对象的API
        //POIFSFileSystem fs = new POIFSFileSystem(is);
        XSSFWorkbook workbook = new XSSFWorkbook(is);

        //3、创建数据抽取类
        XSSFExcelExtractor excelExtractor = new XSSFExcelExtractor(workbook);
        excelExtractor.setIncludeSheetNames(false);//不获取Excel工作表名称
        excelExtractor.setIncludeHeadersFooters(false);
        System.out.println("获取到的内容：" + excelExtractor.getText());
    }
}

