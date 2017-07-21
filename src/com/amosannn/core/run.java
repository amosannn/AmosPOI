package com.amosannn.core;

import java.io.BufferedReader;
import java.io.InputStreamReader;

/**
 * @author amos.lin
 *
 */
public class run {

    /**
     * @param args
     */
    public static void main(final String[] args) {
        try {
            final BufferedReader br = new BufferedReader(new InputStreamReader(System.in));
            final String projectPath = System.getProperty("user.dir");
            System.out.println("文件导入与生成的根路径为项目根目录。");
            System.out.println("请输入待转换的文件名（需含后缀.xls或.xlsx）");
            final String filePath = projectPath + "\\" + br.readLine();
            System.out.println("请输入输出的txt文件文件名（不含后缀名）");
            final String outputTXTPath = projectPath + "\\" + br.readLine()+".txt";
            System.out.println("请输入输出的xml文件文件名（不含后缀名）");
            final String outputXMLPath = projectPath + "\\" + br.readLine()+".xml";
            System.out.println("........................");
            System.out.println("......正在生成文件......");
            System.out.println("........................");
            System.out.println("输入文件    " + filePath + "\n");
            TXTParser.toTXT(filePath, outputTXTPath);// 生成TXT文件
            System.out.println("输出文件    " + outputTXTPath);
            XMLParser.toXML(filePath, outputXMLPath);
            System.out.println("输出文件    " + outputXMLPath);

            System.out.println("........................");
            System.out.println("文件转换成功！");

        } catch (final Exception e) {
            e.printStackTrace();
        }
    }

}
