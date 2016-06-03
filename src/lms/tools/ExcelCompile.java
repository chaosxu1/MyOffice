/**
 * 
 */
package lms.tools;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import lms.data.BomItem;

/**
 * @author xusu1
 *
 */
public class ExcelCompile {
	public static String filename = null;
	private static String jarpach;
	public ExcelCompile() {
		super();
		
		  String path = Thread.currentThread().getContextClassLoader().getResource("").getPath();    //添加
		    ;
	        
		// TODO Auto-generated constructor stub
	}

	// private static String xls2003 = "product2.xlsx";
	private static String xlsx2007 = "product2.xlsx";
	private static String samplefilename="sample.xlsx";
	private static   List<BomItem>    bomItemList = new ArrayList<BomItem>();// 返回封装数据的List ;
    public static String workpath="D:";
	/**
	 * @param args1
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		ExcelCompile  ec= new ExcelCompile();
		
		File excelFile = null;// Excel文件对象
		 
		excelFile = new File(xlsx2007);  
		ec.readFromXLS2003(excelFile);
		System.out.println("bomItemList.size()" + bomItemList.size() );
		 ec.formatOut();
		 
		 
		
		 
		 
	}

	
	/**
	 * Highlight cells based on their values
	 */
	static void createCell(Sheet sheet  ) {

	   List  result = new ArrayList();
		BomItem bomitem;

	int rowNum=10;
		 
		for (int ii=0;ii<bomItemList.size();ii++){
			
			bomitem=	(BomItem) bomItemList.get(ii);
			String desc=bomitem.getFeatureCodeDesc();
			
			 if(bomitem.getBomCode().contains("System x")){
				 rowNum = rowNum + 1;
					Row row;
					sheet.createRow(rowNum++);
					if (sheet.getRow(rowNum) == null)
						row = sheet.createRow(rowNum);
					else
						row = sheet.getRow(rowNum);
					row.createCell(1)
							.setCellValue(bomitem.getBomCode());
					
				
					 
					continue;
			 }
			 
			 
			 if(bomitem.getBomCode().contains("BOMCode")){
				 rowNum = rowNum + 1;
					Row row;
					if (sheet.getRow(rowNum) == null)
						row = sheet.createRow(rowNum);
					else
						row = sheet.getRow(rowNum);
					row.createCell(1)
							.setCellValue( "");
					
					 
					continue;
			 }
			 
			 if(!desc.isEmpty()&& desc!=null){
				
				 
				 
				if (desc.contains("Intel Xeon Processor")
						|| desc.contains("TruDDR4 Memory")
						||desc.contains("RDIMM")
						|| desc.contains("ServeRAID")
						|| desc.contains("G3HS")
						|| desc.contains("PCIe Adapter")
						||desc.contains("GbE Adapter")
						||desc.contains("SFP+ Adapter")
						|| (desc.contains("Emulex") && !desc.contains("2U Bracke"))
						||desc.contains("QLogic")
						||desc.contains("SR Transceiver")
						
						||desc.contains("PCIe Riser")
						||(desc.contains("Power Supply") && ! desc.contains("Blank Filler"))
						|| desc.contains("Integrated Management Module Advanced")
						||desc.contains("DVD")
						||desc.contains("Line cord")
						||desc.contains("GbE Adapter")
						||desc.contains("Lightpath")
//						|| desc.contains("Power Cable")
						|| desc.contains("Documentation")
						||desc.contains("Slides Kit")
						
						) {

					if (desc.contains("Placement"))
						continue;
					else {
						rowNum = rowNum + 1;
						Row row;
						if (sheet.getRow(rowNum) == null)
							row = sheet.createRow(rowNum);
						else
							row = sheet.getRow(rowNum);
						row.createCell(0)
								.setCellValue(bomitem.getFeatureCode());
						row.createCell(1).setCellValue(desc);
						row.createCell(3).setCellValue(bomitem.getAmount());

					}
						 
				 }
				
			 }
				 
		
		}// end search list
		
	 
	}
	
	public   void formatOut(){
		
		
		File excelFile = null;// Excel文件对象
		 
		excelFile = new File(jarpach+samplefilename);  
	 
		XSSFSheet sheet = null; 
		
			
			// HSSFWorkbook workbook = new HSSFWorkbook(is);// 创建Excel2003文件对象
			XSSFWorkbook workbook = null;
			try {
				
				
				
				  URL url = this.getClass().getResource("/sample.xlsx");
		 
			
				  InputStream is =   url.openStream();
//				  new FileInputStream(excelFile);// 获取文件输入流
				workbook = new XSSFWorkbook(is);
				  sheet = workbook.getSheetAt(0);// 取出第一个工作表，索引是0
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}// 创建Excel2003文件对象
			// HSSFSheet sheet = workbook.getSheetAt(0);// 取出第一个工作表，索引是0
			
			createCell(sheet);
		
			
			// Write the output to a file
			 String temp_str="";  
			    Date dt = new Date();  
			    //最后的aa表示“上午”或“下午”    HH表示24小时制    如果换成hh表示12小时制  
			    SimpleDateFormat sdf = new SimpleDateFormat("DDHHmmss");  
			    temp_str=sdf.format(dt);  
			    
			    String file = workpath+ "/out"+temp_str+".xls";
			    ExcelCompile.filename=file;
			
			if (workbook instanceof XSSFWorkbook)
				file += "x";
			FileOutputStream out;
			try {
				out = new FileOutputStream(file);
				workbook.write(out);
				out.close();
				System.out.println("Generated: " + file);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
	
	}
	
	public   void readFromXLS2003(File excelFile) {
		 
		InputStream is = null;// 输入流对象
		String cellStr = null;// 单元格，最终按字符串处理
	  
	
		
		BomItem bomitem = new BomItem();// 每一个bomitem信息对象
		Workbook workbook;
		Sheet sheet = null;
		 
		try {
			 
			is = new FileInputStream(excelFile);// 获取文件输入流
			// HSSFWorkbook workbook = new HSSFWorkbook(is);// 创建Excel2003文件对象
			   workbook = new XSSFWorkbook(is);// 创建Excel2003文件对象
			  sheet = workbook.getSheetAt(0);// 取出第一个工作表，索引是0
		
		} catch (OLE2NotOfficeXmlFileException e) {
			 
			try {
				is = new FileInputStream(excelFile);// 获取文件输入流
				workbook = new HSSFWorkbook(is);
				sheet = workbook.getSheetAt(0);// 取出第一个工作表，索引是0
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}// 创建Excel2003文件对象
			   
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
		
		
		finally {// 关闭文件流
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		
		
		
	
		// 开始循环遍历行，表头不处理，从1开始

		
		for (int i = 0; i <= sheet.getLastRowNum(); i++) {

			// HSSFRow row = sheet.getRow(i);// 获取行对象
			 Row row = sheet.getRow(i);// 获取行对象

			if (row == null) {// 如果为空，不处理
				continue;
			}
			// 循环遍历单元格
			BomItem bomitem1 = new BomItem();// 实例化bomitem对象
			for (int j = 0; j < row.getLastCellNum(); j++) {
				
				// HSSFCell cell = row.getCell(j);// 获取单元格对象
				 Cell cell = row.getCell(j);// 获取单元格对象

				if (cell == null) {// 单元格为空设置cellStr为空串
					cellStr = "";
				} else if (cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {// 对布尔值的处理
					cellStr = String.valueOf(cell.getBooleanCellValue());
				} else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {// 对数字值的处理
					cellStr =(int) cell.getNumericCellValue() + "";
				} else {// 其余按照字符串处理
					cellStr = cell.getStringCellValue();
				}


					bomitem1.setNumber(""+i);
					if (j == 1) {
						bomitem1.setBomCode(cellStr);

					}
					if (j == 2) {
						bomitem1.setFeatureCodeDesc(cellStr);
					}
					if (j == 5) {
						bomitem1.setFeatureCode(cellStr);
					}
					if (j == 6) {
						bomitem1.setAmount(cellStr);
					}
//				}// end of if
			}// end of shell
		 
			bomItemList.add(bomitem1);// 数据装入List

		} // end of row
		
		
		
		System.out.print(bomItemList);
		 
	}

	 
}
