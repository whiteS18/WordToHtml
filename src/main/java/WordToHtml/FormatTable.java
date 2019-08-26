package WordToHtml;


import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * @Description: docx文件转HTML后表格格式会混乱，统一表格格式
 * Param:
 * @return:
 * @Author: CW Song
 * @Date: 2019/8/26
 */
public class FormatTable {
    static StringBuilder sb = new StringBuilder();
    static String str;

    public static void main(String[] args) {
        BufferedReader br;
        ArrayList<String> list = new ArrayList<>();
        try {
            //这里参数是你需要转化的html
            br = new BufferedReader(new FileReader("C:/demo/aaa.html"));
            while ((str = br.readLine()) != null) // 判断最后一行不存在，为空结束循环
            {
                // System.out.println(str);//原样输出读到的内容
                sb.append(str);
                sb.append("\r\n");
            }
            br.close();
            String s = sb.toString();
            ReadWordTable readWordTable = new ReadWordTable();
            List<XWPFTable> wordTables = readWordTable.getWordTables();
            String s1 = s.replaceAll("<table", "<table1><table");
            String s2 = s1.replaceAll("</table>", "</table><table2>");
            String[] split = s2.split("<table1>");
            String[] split1;
            int i = 0;
            StringBuffer str5 = new StringBuffer();
            for (String sp : split) {
                split1 = sp.split("<table2>");
                for (String sp1 : split1) {
                    if (sp1.contains("<table")) {
                        sp1 = readWordTable.readTable(wordTables.get(i));
                        System.out.println(i);
                        i++;
                    }
                    str5.append(sp1);
                }
            }
            //这里参数是输出到那里，也是传全路径
            BufferedWriter out = new BufferedWriter(new FileWriter("C:/demo/aaa.html"));
            out.write(str5.toString());
            out.newLine();
            out.flush();
            out.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

}
