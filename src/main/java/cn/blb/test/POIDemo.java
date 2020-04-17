package cn.blb.test;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

public class POIDemo {
    public static void main(String[] args) throws IOException {
        /**
         * poi函数库中：
         *      工作簿：XSSWorkbook
         *      工作表：XSSFSheet
         *      行和单元格：Row / Cell
         */
        //获取一个Excel表中的数据
        //1、获取excle工作簿对象
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook("c:\\Users\\Administrator.KQ3XM23MUKVL0OJ\\Desktop\\TestPOI.xlsx");

        //2、获取工作表
        XSSFSheet xssfSheet = xssfWorkbook.getSheet("name");
        //xssfWorkbook.getSheetAt(0);//通过下标获取第一个工作表

        //3、获取行对象
        for(Row row : xssfSheet){
            //4、获取列对象(当前行中的单元格中的内容)
            for(Cell cell : row){
                //输出单元格中的内容
                System.out.println(cell.getStringCellValue());
                cell.setCellValue("哈哈");
                System.out.println(cell.getStringCellValue());
            }
        }

    }
}
