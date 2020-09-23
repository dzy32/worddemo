package com.example.worddemo;

import com.example.worddemo.util.WordUtil;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDPt;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @author ys
 * @date 2020/9/1 16:15
 */
public class ChartTest {
    // Methode to set title in the data sheet without creating a Table but using the sheet data only.
    // Creating a Table is not really necessary.
//    static CellReference setTitleInDataSheet(XWPFChart chart, String title, int column) throws Exception {
//        XSSFWorkbook workbook = chart.getWorkbook();
//        XSSFSheet sheet = workbook.getSheetAt(0);
//        XSSFRow row = sheet.getRow(0);
//        if (row == null)
//            row = sheet.createRow(0);
//        XSSFCell cell = row.getCell(column);
//        if (cell == null)
//            cell = row.createCell(column);
//        cell.setCellValue(title);
//        return new CellReference(sheet.getSheetName(), 0, column, true, true);
//    }

    public static void main(String[] args) throws Exception {
//        try (XWPFDocument document = new XWPFDocument()) {
//
//            // create the data
//            String[] categories = new String[] { "Lang 1", "Lang 2", "Lang 3" };
//            Double[] valuesA = new Double[] { 10d, 20d, 40d };
//            Double[] valuesB = new Double[] { 10d, 20d, 30d };
//
//            // create the chart
//            XWPFChart chart = document.createChart(15 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);
//
//            // create data sources
//            int numOfPoints = categories.length;
//            String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
//            String valuesDataRangeA = chart.formatRange(new CellRangeAddress(1, numOfPoints, 1, 1));
//            String valuesDataRangeB = chart.formatRange(new CellRangeAddress(1, numOfPoints, 2, 2));
//            XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);
//            XDDFNumericalDataSource<Double> valuesDataA = XDDFDataSourcesFactory.fromArray(valuesA, valuesDataRangeA, 1);
//            XDDFNumericalDataSource<Double> valuesDataB = XDDFDataSourcesFactory.fromArray(valuesB, valuesDataRangeB, 2);
//
//            // create axis
//            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
//            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
//            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
//            // Set AxisCrossBetween, so the left axis crosses the category axis between the categories.
//            // Else first and last category is exactly on cross points and the bars are only half visible.
//            leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
//
//            // create chart data
//            XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
//            ((XDDFBarChartData) data).setBarDirection(BarDirection.COL);
//
//            // create series
//            // if only one series do not vary colors for each bar
//            // 设置每一列的数据颜色
//            ((XDDFBarChartData) data).setVaryColors(true);
//            XDDFChartData.Series series = data.addSeries(categoriesData, valuesDataA);
//            // XDDFChart.setSheetTitle is buggy. It creates a Table but only half way and incomplete.
//            // Excel cannot opening the workbook after creatingg that incomplete Table.
//            // So updating the chart data in Word is not possible.
//            //series.setTitle("a", chart.setSheetTitle("a", 1));
//            series.setTitle("a", setTitleInDataSheet(chart, "a", 1));
////            series.setTitle("b",setTitleInDataSheet(chart,"b",2));
//
//
//			/*
//			   // if more than one series do vary colors of the series
//			   ((XDDFBarChartData)data).setVaryColors(true);
//			   series = data.addSeries(categoriesData, valuesDataB);
//			   //series.setTitle("b", chart.setSheetTitle("b", 2));
//			   series.setTitle("b", setTitleInDataSheet(chart, "b", 2));
//			*/
//            CTPlotArea plotArea = chart.getCTChart().getPlotArea();
//            // plot chart data
//            chart.plot(data);
//            //给每个条形图设置颜色
//            CTBarSer ser = plotArea.getBarChartArray(0).getSerArray(0);
//            //黑色
//            CTDPt dpt = ser.addNewDPt();
//            dpt.addNewIdx().setVal(0);
//            dpt.addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(new byte[] { (byte) 0, (byte) 0, (byte) 0 });
//            //红色
//            dpt = ser.addNewDPt();
//            dpt.addNewIdx().setVal(1);
//            dpt.addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(new byte[] { (byte) 255, (byte) 0, (byte) 0 });
//            //红色
//            dpt = ser.addNewDPt();
//            dpt.addNewIdx().setVal(2);
//            dpt.addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(new byte[] { (byte) 255, (byte) 0, (byte) 255 });
//            // create legend
//            XDDFChartLegend legend = chart.getOrAddLegend();
//            legend.setPosition(LegendPosition.RIGHT);
//            legend.setOverlay(false);

            // Write the output to a file
            XWPFDocument document = new XWPFDocument();
            String[] categories = new String[] {"第一","second","third"};
            Integer[] values = new Integer[] {20, 30, 90};
            List<byte[]> colors = new ArrayList<>();
            colors.add(new byte[] { (byte) 0, (byte) 0, (byte) 0 });
            colors.add(new byte[] { (byte) 255, (byte) 0, (byte) 0 });
            colors.add(new byte[] { (byte) 255, (byte) 0, (byte) 255 });
            WordUtil.addChart(values,categories,colors,"测试例子",document);
            try (FileOutputStream fileOut = new FileOutputStream("E:\\chromeDownload\\c.docx")) {
                document.write(fileOut);
            }catch(IOException e){
                e.printStackTrace();
            }

    }

}
