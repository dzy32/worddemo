package com.example.worddemo.util;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.List;
import java.util.Map;

/**
 * Word工具类
 * @author liuyuay
 * */
@Component
public class WordUtil1 {

    private static final Logger LOGGER = LoggerFactory.getLogger(WordUtil.class);

    /**
     * 设置一级标题内容及样式
     * @param document 文本对象
     * @param paragraph 段落
     * @param text 标题内容
     */
    public static void setLevelTitleFirst(XWPFDocument document, XWPFParagraph paragraph, String text) {
        // 将段落原有文本(原有所有的Run)全部删除
        deleteRun(paragraph);
        // 插入新的Run即将新的文本插入段落
        var createRun = paragraph.insertNewRun(0);
        createRun.setText(text);
        createRun.setFontSize(16);
        createRun.setFontFamily("黑体");
        addCustomHeadingStyle(document, "标题1", 1);
        paragraph.setStyle("标题1");
    }

    /**
     * 设置二级标题内容及样式
     * @param document 文本对象
     * @param paragraph 段落
     * @param text 标题内容
     */
    public static void setLevelTitleSecond(XWPFDocument document, XWPFParagraph paragraph, String text) {
        // 将段落原有文本(原有所有的Run)全部删除
        deleteRun(paragraph);
        // 插入新的Run即将新的文本插入段落
        var createRun = paragraph.insertNewRun(0);
        // 设置段落文本
        createRun.setText(text);
        // 设置字体大小
        createRun.setFontSize(16);
        // 是否粗体
        createRun.setBold(true);
        // 设置字体
        createRun.setFontFamily("楷体_GB2312");
        // 下面三行代码可使标题缩进
        paragraph.setIndentationFirstLine(600);
        paragraph.setSpacingAfter(10);
        paragraph.setSpacingBefore(10);

        addCustomHeadingStyle(document, "标题2", 2);
        paragraph.setStyle("标题2");
    }

    /**
     * 设置三级标题内容及样式
     * @param document 文本对象
     * @param paragraph 段落
     * @param text 标题内容
     */
    public static void setLevelTitleThird(XWPFDocument document, XWPFParagraph paragraph, String text) {
        // 将段落原有文本(原有所有的Run)全部删除
        deleteRun(paragraph);
        // 插入新的Run即将新的文本插入段落
        var createRun = paragraph.insertNewRun(0);
        // 设置段落文本
        createRun.setText(text);
        // 设置字体大小
        createRun.setFontSize(16);
        // 是否粗体
        createRun.setBold(true);
        // 设置字体
        createRun.setFontFamily("楷体_GB2312");

        addCustomHeadingStyle(document, "标题3", 3);
        paragraph.setStyle("标题3");
    }

    /**
     * 将段落原有文本全部删除
     * 在创建标题时有调用
     * @param paragraph 段落对象
     * */
    public static void deleteRun(XWPFParagraph paragraph) {
        // 将段落原有文本(原有所有的Run)全部删除
        List<XWPFRun> runs = paragraph.getRuns();
        var runSize = runs.size();
        // Paragrap中每删除一个run,其所有的run对象就会动态变化，即不能同时遍历和删除
        var haveRemoved = 0;
        for (var runIndex = 0; runIndex < runSize; runIndex++) {
            paragraph.removeRun(runIndex - haveRemoved);
            haveRemoved++;
        }
    }

    /**
     * 设置四级标题内容及样式
     * @param document 文本对象
     * @param paragraph 段落
     * @param text 标题内容
     */
    public static void setLevelTitleFourth(XWPFDocument document, XWPFParagraph paragraph, String text) {
        // 将段落原有文本(原有所有的Run)全部删除
        deleteRun(paragraph);
        // 插入新的Run即将新的文本插入段落
        var createRun = paragraph.insertNewRun(0);
        // 设置段落文本
        createRun.setText(text);
        // 设置字体大小
        createRun.setFontSize(16);
        // 是否粗体
        createRun.setBold(true);
        // 设置字体
        createRun.setFontFamily("楷体_GB2312");
        // 下面三行代码可使标题缩进
        paragraph.setIndentationFirstLine(600);
        paragraph.setSpacingAfter(10);
        paragraph.setSpacingBefore(10);

        addCustomHeadingStyle(document, "标题4", 4);
        paragraph.setStyle("标题4");
    }

    /**
     * 增加自定义标题样式。这里用的是stackoverflow的源码
     * @param docxDocument 目标文档
     * @param strStyleId   样式名称
     * @param headingLevel 样式级别
     */
    private static void addCustomHeadingStyle(XWPFDocument docxDocument, String strStyleId, int headingLevel) {
        var ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(strStyleId);

        var styleName = CTString.Factory.newInstance();
        styleName.setVal(strStyleId);
        ctStyle.setName(styleName);

        var indentNumber = CTDecimalNumber.Factory.newInstance();
        indentNumber.setVal(BigInteger.valueOf(headingLevel));

        // lower number > style is more prominent in the formats bar
        ctStyle.setUiPriority(indentNumber);

        var onoffnull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onoffnull);

        // style shows up in the formats bar
        ctStyle.setQFormat(onoffnull);

        // style defines a heading of the given level
        var ppr = CTPPr.Factory.newInstance();
        ppr.setOutlineLvl(indentNumber);
        ctStyle.setPPr(ppr);

        var style = new XWPFStyle(ctStyle);

        // is a null op if already defined
        var styles = docxDocument.createStyles();

        style.setType(STStyleType.PARAGRAPH);
        styles.addStyle(style);

    }

    /**
     * 添加第一页，生成漏洞整改报告报表时有用上
     * @param docx 文本对象
     * @param text 内容
     * */
    public static void addHeaderPage(XWPFDocument docx, String text){
        // 标题
        var titleParagraph = docx.createParagraph();
        // 设置居中
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);
        // 创建run(用于向段落中添加文本)
        var run = titleParagraph.createRun();
        // 回车
//        run.addCarriageReturn();
//        run.addCarriageReturn();
//        run.addCarriageReturn();
        // 设置文本
        run.setText(text);
        run.addCarriageReturn();
        run.setText("目录");
        run.addCarriageReturn();
        var flag = docx.createParagraph();
        var run1 = flag.createRun();
        run1.setText("齉");
        run.setFontSize(25);
        run.setBold(true);
        // 强制分页
//        run1.addBreak(BreakType.PAGE);
    }

    /**
     * 插入图片
     * @param docx word文本对象
     * @param imageAddr 图片地址
     * @param width 图片宽度
     * @param height 高度
     * */
    public static void insertImage(XWPFDocument docx,String imageAddr,int width,int height) throws IOException, InvalidFormatException {
        // 创建段落插入图片
        var pictureOne = docx.createParagraph();
        var pictureOneRun = pictureOne.createRun();

        // 获取图片后缀
        var prefix = imageAddr.substring(imageAddr.lastIndexOf(".")+1);

        var pictureType = 0;

        switch (prefix) {
            case "png":
                pictureType = XWPFDocument.PICTURE_TYPE_PNG;
                break;
            case "jpg":
            case "jpeg":
                pictureType = XWPFDocument.PICTURE_TYPE_JPEG;
                break;
            case "gif":
                pictureType = XWPFDocument.PICTURE_TYPE_GIF;
                break;
            default:
                System.out.println("图片格式错误！");
                return;
        }

        FileInputStream is;
        try {
            // 如果将图片放在jar包中来使用，这里必须转成输入流，否则无法读取图片
            is = new FileInputStream(imageAddr);
            pictureOneRun.addCarriageReturn();
            pictureOneRun.addPicture(is, pictureType, imageAddr, Units.toEMU(width), Units.toEMU(height));
            is.close();
        } catch (FileNotFoundException e) {
            LOGGER.error("图片未找到:"+imageAddr);
        }
    }

    /**
     * 生成目录
     * @param document word文本对象
     * @param flag 目录生成的标记位置 （即目录在该字符串出现的位置生成， 要求在Word文档中该字符串只能出现一次）
     */
    public static void generateTOC(XWPFDocument document, String flag) {
        String replaceText = "";
        for (XWPFParagraph p : document.getParagraphs()) {
            for (XWPFRun r : p.getRuns()) {
                int pos = r.getTextPosition();
                String text = r.getText(pos);
                if (text != null && text.contains(flag)) {
                    text = text.replace(flag, replaceText);
                    r.setText(text, 0);
                    // 这个设置根据Office的官方文档来设置
                    addField(p, "TOC \\o \"1-5\" \\h \\z \\u");
                    break;
                }
            }
        }
    }

    /**
     * 添加目录
     * @param paragraph 段落对象
     * @param fieldName 抓取什么字段来判断是否需要生成目录
     * */
    private static void addField(XWPFParagraph paragraph, String fieldName) {
        CTSimpleField ctSimpleField = paragraph.getCTP().addNewFldSimple();
        ctSimpleField.setInstr(fieldName);
        ctSimpleField.setDirty(STOnOff.TRUE);
        ctSimpleField.addNewR().addNewT().setStringValue("<<fieldName>>");
    }

    /**
     * 替换柱状图数据
     */
    public static void replaceBarCharts(XWPFChart chart, List<String> titleArr,
                                 List<String> fldNameArr, List<Map<String, String>> listItemsByType) {
        chart.getCTChart();

        //根据属性第一列名称切换数据类型
        CTChart ctChart = chart.getCTChart();
        CTPlotArea plotArea = ctChart.getPlotArea();

        CTBarChart barChart = plotArea.getBarChartArray(0);
        List<CTBarSer> BarSerList = barChart.getSerList();  // 获取柱状图单位

        //刷新内置excel数据
        refreshExcel(chart, listItemsByType, fldNameArr, titleArr);
        //刷新页面显示数据
        refreshBarStrGraphContent(barChart, BarSerList, listItemsByType, fldNameArr, 1);
    }

    /**
     * 刷新内置excel数据
     *
     * @param chart
     * @param dataList
     * @param fldNameArr
     * @param titleArr
     * @return
     */
    public static boolean refreshExcel(XWPFChart chart,
                                List<Map<String, String>> dataList, List<String> fldNameArr, List<String> titleArr) {
        boolean result = true;
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");
        //根据数据创建excel第一行标题行
        for (int i = 0; i < titleArr.size(); i++) {
            if (sheet.getRow(0) == null) {
                sheet.createRow(0).createCell(i).setCellValue(titleArr.get(i) == null ? "" : titleArr.get(i));
            } else {
                sheet.getRow(0).createCell(i).setCellValue(titleArr.get(i) == null ? "" : titleArr.get(i));
            }
        }

        //遍历数据行
        for (int i = 0; i < dataList.size(); i++) {
            Map<String, String> baseFormMap = dataList.get(i);//数据行
            //fldNameArr字段属性
            for (int j = 0; j < fldNameArr.size(); j++) {
                if (sheet.getRow(i + 1) == null) {
                    if (j == 0) {
                        try {
                            sheet.createRow(i + 1).createCell(j).setCellValue(baseFormMap.get(fldNameArr.get(j)) == null ? "" : baseFormMap.get(fldNameArr.get(j)));
                        } catch (Exception e) {
                            if (baseFormMap.get(fldNameArr.get(j)) == null) {
                                sheet.createRow(i + 1).createCell(j).setCellValue("");
                            } else {
                                sheet.createRow(i + 1).createCell(j).setCellValue(baseFormMap.get(fldNameArr.get(j)));
                            }
                        }
                    }
                } else {
                    BigDecimal b = new BigDecimal(baseFormMap.get(fldNameArr.get(j)));
                    double value = 0d;
                    if (b != null) {
                        value = b.doubleValue();
                    }
                    if (value == 0) {
                        sheet.getRow(i + 1).createCell(j);
                    } else {
                        sheet.getRow(i + 1).createCell(j).setCellValue(b.doubleValue());
                    }
                }
            }

        }
        // 更新嵌入的workbook
        POIXMLDocumentPart xlsPart = chart.getRelations().get(0);
        OutputStream xlsOut = xlsPart.getPackagePart().getOutputStream();

        try {
            wb.write(xlsOut);
            xlsOut.close();
        } catch (IOException e) {
            e.printStackTrace();
            result = false;
        } finally {
            if (wb != null) {
                try {
                    wb.close();
                } catch (IOException e) {
                    e.printStackTrace();
                    result = false;
                }
            }
        }
        return result;
    }

    /**
     * 刷新柱状图数据方法
     *
     * @param typeChart
     * @param serList
     * @param dataList
     * @param fldNameArr
     * @param position
     * @return
     */
    public static boolean refreshBarStrGraphContent(Object typeChart, List<?> serList, List<Map<String, String>> dataList,
                                             List<String> fldNameArr, int position) {
        boolean result = true;
        //更新数据区域
        for (int i = 0; i < serList.size(); i++) {
//            CTSerTx tx=null;
            CTAxDataSource cat = null;
            CTNumDataSource val = null;
            CTBarSer ser = ((CTBarChart) typeChart).getSerArray(i);
//            tx= ser.getTx();

            // Category Axis Data
            cat = ser.getCat();
            // 获取图表的值
            val = ser.getVal();
            // strData.set
            CTStrData strData = cat.getStrRef().getStrCache();
            CTNumData numData = val.getNumRef().getNumCache();
            // unset old axis text
            strData.setPtArray((CTStrVal[]) null);
            // unset old values
            numData.setPtArray((CTNumVal[]) null);

            // set model
            long idx = 0;
            for (int j = 0; j < dataList.size(); j++) {
                //判断获取的值是否为空
                String value = "0";
                if (new BigDecimal(dataList.get(j).get(fldNameArr.get(i + position))) != null) {
                    value = new BigDecimal(dataList.get(j).get(fldNameArr.get(i + position))).toString();
                }
                if (!"0".equals(value)) {
                    CTNumVal numVal = numData.addNewPt();//序列值
                    numVal.setIdx(idx);
                    numVal.setV(value);
                }
                CTStrVal sVal = strData.addNewPt();//序列名称
                sVal.setIdx(idx);
                sVal.setV(dataList.get(j).get(fldNameArr.get(0)));
                idx++;
            }
            numData.getPtCount().setVal(idx);
            strData.getPtCount().setVal(idx);


            //赋值横坐标数据区域
            String axisDataRange = new CellRangeAddress(1, dataList.size(), 0, 0)
                    .formatAsString("Sheet1", true);
            cat.getStrRef().setF(axisDataRange);

            //数据区域
            String numDataRange = new CellRangeAddress(1, dataList.size(), i + position, i + position)
                    .formatAsString("Sheet1", true);
            val.getNumRef().setF(numDataRange);

        }
        return result;
    }
}
