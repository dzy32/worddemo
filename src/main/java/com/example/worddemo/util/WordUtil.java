package com.example.worddemo.util;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.LocaleUtil;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDPt;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;

/**
 * Word工具类
 * */
@Component
public class WordUtil {

    private static final Logger LOGGER = LoggerFactory.getLogger(WordUtil.class);

    /**
     * 设置一级标题内容及样式
     *
     * @param document  文本对象
     * @param paragraph 段落
     * @param text      标题内容
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
     *
     * @param document  文本对象
     * @param paragraph 段落
     * @param text      标题内容
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
     * 添加第一页
     *
     * @param docx 文本对象
     * @param text 内容
     */
    public static void addHeaderPage(XWPFDocument docx, String text) {
        // 标题
        var titleParagraph = docx.createParagraph();
        // 设置居中
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);
        // 创建run(用于向段落中添加文本)
        var run = titleParagraph.createRun();
        // 换行
        // run.addBreak();
        // 回车
        run.addCarriageReturn();
        run.addCarriageReturn();
        run.addCarriageReturn();
        // 设置文本
        run.setText(text);
        run.setFontSize(25);
        run.setBold(true);
        // 强制分页
        run.addBreak(BreakType.PAGE);
    }

    /**
     * 插入图片
     */
    public static void insertImage(XWPFDocument docx, String imageAddr) throws IOException, InvalidFormatException {
        // 创建段落插入图片
        var pictureOne = docx.createParagraph();
        var pictureOneRun = pictureOne.createRun();

        FileInputStream is = null;
        try {
            is = new FileInputStream(imageAddr);
            pictureOneRun.addCarriageReturn();
            // 420x300 pixels
            pictureOneRun.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, imageAddr, Units.toEMU(430), Units.toEMU(300));
            is.close();
        } catch (FileNotFoundException e) {
            LOGGER.error("图片未找到:" + imageAddr);
        }
    }

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
     * 增加自定义标题样式。这里用的是stackoverflow的源码
     *
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
     * 增加柱形图
     * @param values 数据
     * @param categories 每一列的名称
     * @param document
     * @param title 图表标题
     * @throws Exception
     */
    public static void addChart(Integer[] values, String[] categories, List<byte[]> colors, String title,XWPFDocument document) throws Exception{
        // create the data

//        Double[] valuesA = new Double[] { 10d, 20d, 30d };
        //目前不了解 valuesb的作用
//        Integer[] valuesB = new Integer[]{ 15, 25, 35 };
        // create the chart
        XWPFChart chart = document.createChart(15 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);

        // create data sources
        int numOfPoints = categories.length;
        String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
        String valuesDataRangeA = chart.formatRange(new CellRangeAddress(1, numOfPoints, 1, 1));
        String valuesDataRangeB = chart.formatRange(new CellRangeAddress(1, numOfPoints, 2, 2));
        XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);
        XDDFNumericalDataSource<Integer> valuesDataA = XDDFDataSourcesFactory.fromArray(values, valuesDataRangeA, 1);
//        XDDFNumericalDataSource<Integer> valuesDataB = XDDFDataSourcesFactory.fromArray(valuesB, valuesDataRangeB, 2);

        // create axis
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        // Set AxisCrossBetween, so the left axis crosses the category axis between the categories.
        // Else first and last category is exactly on cross points and the bars are only half visible.
        leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

        // create chart data
        XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
        ((XDDFBarChartData) data).setBarDirection(BarDirection.COL);

        // create series
        // if only one series do not vary colors for each bar
        // 设置每一列的数据颜色，设置为true才可以自定义颜色
        ((XDDFBarChartData) data).setVaryColors(true);
        XDDFChartData.Series series = data.addSeries(categoriesData, valuesDataA);
        // XDDFChart.setSheetTitle is buggy. It creates a Table but only half way and incomplete.
        // Excel cannot opening the workbook after creatingg that incomplete Table.
        // So updating the chart data in Word is not possible.
        //series.setTitle("a", chart.setSheetTitle("a", 1));
        series.setTitle(title, setTitleInDataSheet(chart, title, 1));
//            series.setTitle("b",setTitleInDataSheet(chart,"b",2));


			/*
			   // if more than one series do vary colors of the series
			   ((XDDFBarChartData)data).setVaryColors(true);
			   series = data.addSeries(categoriesData, valuesDataB);
			   //series.setTitle("b", chart.setSheetTitle("b", 2));
			   series.setTitle("b", setTitleInDataSheet(chart, "b", 2));
			*/

        // plot chart data 绘制图表
        chart.plot(data);
        CTPlotArea plotArea = chart.getCTChart().getPlotArea();
        // plot chart data
        chart.plot(data);
        //给每个条形图设置颜色
        CTBarSer ser = plotArea.getBarChartArray(0).getSerArray(0);
        CTDPt dpt = ser.addNewDPt();
        for(int i =0;i<colors.size();i++){
            if(i != 0) {
                dpt = ser.addNewDPt();
            }
            dpt.addNewIdx().setVal(i);
            dpt.addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(colors.get(i));
        }
        // create legend
        XDDFChartLegend legend = chart.getOrAddLegend();
        //设置小标题的位置
        legend.setPosition(LegendPosition.RIGHT);
        legend.setOverlay(false);
    }

    // Methode to set title in the data sheet without creating a Table but using the sheet data only.
    // Creating a Table is not really necessary.
    private static CellReference setTitleInDataSheet(XWPFChart chart, String title, int column) throws Exception {
        XSSFWorkbook workbook = chart.getWorkbook();
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow row = sheet.getRow(0);
        if (row == null)
            row = sheet.createRow(0);
        XSSFCell cell = row.getCell(column);
        if (cell == null)
            cell = row.createCell(column);
        cell.setCellValue(title);
        return new CellReference(sheet.getSheetName(), 0, column, true, true);
    }

    /**
     *
     * @param document
     * @param flag 目录生成的标记位置 （即目录在该字符串出现的位置生成， 要求在Word文档中该字符串只能出现一次）
     * @throws InvalidFormatException
     * @throws FileNotFoundException
     * @throws IOException
     */
    public static void generateTOC(XWPFDocument document, String flag) throws InvalidFormatException, FileNotFoundException, IOException {
        String findText = flag;
        String replaceText = "";
        for (XWPFParagraph p : document.getParagraphs()) {
            for (XWPFRun r : p.getRuns()) {
                int pos = r.getTextPosition();
                String text = r.getText(pos);
//                打印每一段，测试用
                System.out.println(text);
                if (text != null && text.contains(findText)) {
                    text = text.replace(findText, replaceText);
                    r.setText(text, 0);
                    addField(p, "TOC \\o \"1-3\" \\h \\z \\u");
//                    addField(p, "TOC \\h");
                    break;
                }
            }
        }
    }

    private static void addField(XWPFParagraph paragraph, String fieldName) {
        CTSimpleField ctSimpleField = paragraph.getCTP().addNewFldSimple();
        ctSimpleField.setInstr(fieldName);
        ctSimpleField.setDirty(STOnOff.TRUE);
        ctSimpleField.addNewR().addNewT().setStringValue("<<fieldName>>");
    }
}
