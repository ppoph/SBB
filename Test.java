package Excel;





import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Array;
import java.util.Arrays;

public class Test {
    public static void main(String[] args) throws ClassNotFoundException, IOException {
        FileInputStream fip = new FileInputStream("D:\\work\\大数据导论预处理作业.xlsx");//创建输入流
        XSSFWorkbook wb = new XSSFWorkbook(fip);//在输入流中创建工作簿
        XSSFSheet sheet = wb.getSheetAt(0);//在工作簿中获取目标表
        /*XSSFRow row =sheet.getRow(0);
        String s1 = row.getCell(0).getStringCellValue();
        String s2 = row.getCell(1).getStringCellValue();
        String s3 = row.getCell(2).getStringCellValue();
        String s4 = row.getCell(3).getStringCellValue();
        String s5 = row.getCell(4).getStringCellValue();
        String s6 = row.getCell(5).getStringCellValue();
        String s7 = row.getCell(6).getStringCellValue();
        String s8 = row.getCell(7).getStringCellValue();
        String s9 = row.getCell(8).getStringCellValue();
        String s10 = row.getCell(9).getStringCellValue();
        String s11 = row.getCell(10).getStringCellValue();
        System.out.println(s1+"\t"+s2+"\t"+s3+"\t"+s4+"\t"+s5+"\t"+s6);*/
        XSSFRow row;
        int n =sheet.getPhysicalNumberOfRows();
        for (int i = 1; i < 190; i++) {
            row = sheet.getRow(i);
            //double a1 = row.getCell(0).getNumericCellValue();
            String a2 =row.getCell(8).getStringCellValue();
            /*System.out.println(a2);*/
            String b = getSusheNumber(a2);
            /*for (int j = 0; j < arr.length; j++) {
                System.out.println((char)arr[j]);
            }*/
            System.out.println(b);
            //System.out.println(a2);
            //System.out.println();
        }
        fip.close();
    }



    public static String getSusheNumber(String a2){
        char[] arr = new char[5];
        int a =5;
        for (int j = a2.length()-1; j >= 0; j--) {

            char a4 = a2.charAt(j);
            /*System.out.println(a4);*/
            char a5 =a4;
            if(a5=='0'||a4=='1'||a4=='2'||a4=='3'||a4=='4'||a4=='5'||a4=='6'||a4=='7'||a4=='8'||a4=='9'){
                arr[a-1]=a5;
                //System.out.print((char)arr[a-1]);
                a--;
            }
        }
            /*for (int j = 0; j < arr.length; j++) {
                //System.out.println(arr[j]);
            }*/
        if (arr[2]=='1'){
            arr[1]='4';
            arr[0]='1';
        } else if (arr[2]=='9') {
            arr[1]='1';
            arr[0]='2';
        }else {
            String c="错误信息";
            return c;
        }
        String result = ("C"+arr[0]+arr[1]+"-"+arr[2]+arr[3]+arr[4]);
        return result;
    }
}
