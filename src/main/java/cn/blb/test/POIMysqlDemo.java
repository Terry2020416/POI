package cn.blb.test;

import cn.blb.entity.Animal;
import org.apache.commons.dbutils.QueryRunner;
import org.apache.commons.dbutils.handlers.BeanListHandler;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.List;

public class POIMysqlDemo {
    public static void main(String[] args) throws Exception {
        readMysqlToExcle();
    }


    public static void readMysqlToExcle() throws Exception {
        //1、获取数据表中的内容
        QueryRunner qr = new QueryRunner();
        List<Animal> animalList = qr.query(
                DriverManager.getConnection("jdbc:mysql://localhost:3306/mybatis?serverTimezone=UTC","root","tayujiu99")
        ,"select * from animal",new BeanListHandler<Animal>(Animal.class));

        for(Animal animal:animalList){
            System.out.println(animal);
        }
        //2、将数据写入到excel中
        createExcelDatas(animalList);
    }

    private static void createExcelDatas(List<Animal> animalList) throws Exception {
        //1、创建一个工作簿
        XSSFWorkbook wb = new XSSFWorkbook();
        //获取输入流对象
        FileOutputStream fos = new FileOutputStream("C:\\Users\\Administrator.KQ3XM23MUKVL0OJ\\Desktop\\数据库查到的Animal.xlsx");
        //2、创建工作表
        XSSFSheet sheet = wb.createSheet("Animal");
        //3、创建行(标题)
        XSSFRow row = sheet.createRow(0);

        //4、创建列（每列的标题）
        row.createCell(0).setCellValue("id");
        row.createCell(1).setCellValue("name");
        row.createCell(2).setCellValue("kind");
        row.createCell(3).setCellValue("age");

        //创建行(内容)
        for(int i = 0 ; i < animalList.size(); i++){
            XSSFRow row1 =  sheet.createRow(i+1);
            row1.createCell(0).setCellValue(animalList.get(i).getId());
            row1.createCell(1).setCellValue(animalList.get(i).getName());
            row1.createCell(2).setCellValue(animalList.get(i).getKind());
            row1.createCell(3).setCellValue(animalList.get(i).getAge());
        }
        //将工作簿输出到执行文件中
        wb.write(fos);
        //释放资源
        fos.close();
        wb.close();
        System.out.println("写入完毕！！！");
    }
}
