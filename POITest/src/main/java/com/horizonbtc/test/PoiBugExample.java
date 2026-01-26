package com.horizonbtc.test;

import java.awt.Rectangle;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.chart.AxisCrossBetween;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.BarDirection;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFObjectShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;


public class PoiBugExample {

	public static void main(String[] args) {
		boolean writeFile = false;
		
		final Logger logger = LogManager.getLogger();
		
		try {
			// This one works, 
			//InputStream inputStream = PoiBugExample.class.getClassLoader().getResourceAsStream("BugTestTemplateOneSheet.pptx");
			InputStream inputStream = PoiBugExample.class.getClassLoader().getResourceAsStream("BugTestTemplateTwoSheets.pptx");
			XMLSlideShow  ppt = new XMLSlideShow(inputStream);
			inputStream.close();
			logger.debug("PowerPoint template - Available slides:");
			for(int h = 0; h < ppt.getSlides().size(); h++){
				XSLFSlide slide =ppt.getSlides().get(h);
				logger.info("Slide index - "+h);
				for(int i = 0; i < slide.getShapes().size(); i++){
					XSLFShape shape = slide.getShapes().get(i); 
					logger.debug("Slide.shape index - "+i);
					if (shape instanceof XSLFObjectShape) {
						XSLFObjectShape xoShape = (XSLFObjectShape) shape;
						logger.info("Slide shape full name: "+xoShape.getFullName());
						logger.info("Slide shape prog id: "+xoShape.getProgId());
					} else
					if(shape instanceof XSLFTextShape) {
						XSLFTextShape tsh = (XSLFTextShape)shape;
						logger.debug("Is an instance of XSLFTextShape...");
						for(int j = 0; j < tsh.getTextParagraphs().size(); j++){
							logger.debug("XSLFTextShape: XSLFTextParagraph index - "+j);
							XSLFTextParagraph p = tsh.getTextParagraphs().get(j) ;
							logger.debug("XSLFTextParagraph has "+p.getTextRuns().size()+" XSLFTextRuns");
							logger.debug(p.getText());
							for(int k = 0; k < p.getTextRuns().size(); k++){
								XSLFTextRun r = p.getTextRuns().get(k); 
								logger.debug("XSLFTextParagraph: XSLFTextRun index - "+k);
	                        	logger.debug("XSLFTextRun raw text: "+r.getRawText());
	                        	
	                        	if (r.getRawText().contains("[CHART]")){
	                        		logger.info("Chart slide, adding chart...");
	                        		r.setText(null);
	                        		insertChart(slide,ppt);
	                        		writeFile = true;
	                        	}
							}
						}
					} else {
						logger.debug("Not an XSLFTextShape.");
						logger.debug("Shape is "+ shape.getShapeName());
					}
				}
			}
			
			  if (writeFile) {
				  Calendar cal = Calendar.getInstance(); 
				  SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd_H_m"); 
				  String lDate = sdf.format(cal.getTime()); 
				  String filename = "powerpoint_bugtest_" + lDate + ".pptx"; 
				  FileOutputStream fileOut = new FileOutputStream(filename);
				  ppt.write(fileOut); fileOut.close(); 
			  }
			 
			ppt.close();
			
		} catch (Exception e) {
			logger.error(e);
		}
		

	}
	 
	 static void insertChart(XSLFSlide slide, XMLSlideShow  ppt) {
		 XSLFChart chart = ppt.createChart();
		   chart.setTitleText("Bug Test Chart");
		   
		   // set axis
		   XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
		   XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
		   leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
		   leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

		   if (bottomAxis.hasNumberFormat()) bottomAxis.setNumberFormat("@");
		   if (leftAxis.hasNumberFormat()) leftAxis.setNumberFormat("#,##0.00");
		   
		   // define chart data for bar chart
		   XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);

		   // add chart categories (x-axis data)
		   String[] categories = new String[] { "Category 1", "Category 2", "Category 3" };
		   String categoryDataRange = chart.formatRange(new org.apache.poi.ss.util.CellRangeAddress(1, categories.length, 0, 0));
		   XDDFCategoryDataSource categoryData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);

		   // add chart values (y-axis data)
		   Double[] values = new Double[] { 10.0, 20.0, 15.0 };
		   String valuesDataRange = chart.formatRange(new org.apache.poi.ss.util.CellRangeAddress(1, values.length, 1, 1));
		   XDDFNumericalDataSource<Double> valueData = XDDFDataSourcesFactory.fromArray(values, valuesDataRange, 1);
		   XDDFBarChartData bar = (XDDFBarChartData) data;
		   bar.setBarDirection(BarDirection.BAR);

		   // add series
		   XDDFChartData.Series series = data.addSeries(categoryData, valueData);
		   series.setTitle("Series 1", chart.setSheetTitle("Series 1", 1));
		   
		   // plot chart
		   chart.plot(data);

		   // set chart dimensions !!Units are EMU (English Metric Units)!!
		   Rectangle chartDimensions = new Rectangle(
		    100*Units.EMU_PER_POINT, 
		    150*Units.EMU_PER_POINT, 
		    400*Units.EMU_PER_POINT, 
		    300*Units.EMU_PER_POINT);
		   // add chart to slide
		   slide.addChart(chart, chartDimensions);

		 
	 }
}
