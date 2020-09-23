package com.example.worddemo;

import com.example.worddemo.util.WordUtil;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;

import java.io.*;

import static com.example.worddemo.util.WordUtil.generateTOC;

/**
 * @author ys
 * @date 2020/9/1 11:31
 */
public class TestWord {

    public static void main(String []args) throws IOException ,InvalidFormatException{
//        XWPFDocument document= new XWPFDocument();
//
//    //Write the Document in file system
//
//    FileOutputStream out = new FileOutputStream(new File("E:\\chromeDownload\\test.docx"));
//
//    XWPFParagraph titleParagraph = document.createParagraph();
//    WordUtil.setLevelTitleFirst(document,titleParagraph,"一级标题");
//    XWPFParagraph titleParagraph2 = document.createParagraph();
//    WordUtil.setLevelTitleSecond(document,titleParagraph2,"二级标题");
//
//        document.write(out);
//        out.close();
        FileInputStream fileInputStream = new FileInputStream("E:\\chromeDownload\\test.docx");
        XWPFDocument doc = new XWPFDocument(fileInputStream);
        generateTOC(doc, "藟");
        OutputStream out = new FileOutputStream("E:\\chromeDownload\\test.docx");
        doc.write(out);
        out.close();

    }

}
