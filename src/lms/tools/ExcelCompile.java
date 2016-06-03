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
		
		  String path = Thread.currentThread().getContextClassLoader().getResource("").getPath();    //���
		    ;
	        
		// TODO Auto-generated constructor stub
	}

	// private static String xls2003 = "product2.xlsx";
	private static String xlsx2007 = "product2.xlsx";
	private static String samplefilename="sample.xlsx";
	private static   List<BomItem>    bomItemList = new ArrayList<BomItem>();// ���ط�װ���ݵ�List ;
    public static String workpath="D:";
	/**
	 * @param args1
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		ExcelCompile  ec= new ExcelCompile();
		
		File excelFile = null;// Excel�ļ�����
		 
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
		
		
		File excelFile = null;// Excel�ļ�����
		 
		excelFile = new File(jarpach+samplefilename);  
	 
		XSSFSheet sheet = null; 
		
			
			// HSSFWorkbook workbook = new HSSFWorkbook(is);// ����Excel2003�ļ�����
			XSSFWorkbook workbook = null;
			try {
				
				
				
				  URL url = this.getClass().getResource("/sample.xlsx");
		 
			
				  InputStream is =   url.openStream();
//				  new FileInputStream(excelFile);// ��ȡ�ļ�������
				workbook = new XSSFWorkbook(is);
				  sheet = workbook.getSheetAt(0);// ȡ����һ��������������0
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}// ����Excel2003�ļ�����
			// HSSFSheet sheet = workbook.getSheetAt(0);// ȡ����һ��������������0
			
			createCell(sheet);
		
			
			// Write the output to a file
			 String temp_str="";  
			    Date dt = new Date();  
			    //����aa��ʾ�����硱�����硱    HH��ʾ24Сʱ��    �������hh��ʾ12Сʱ��  
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
		 
		InputStream is = null;// ����������
		String cellStr = null;// ��Ԫ�����հ��ַ�������
	  
	
		
		BomItem bomitem = new BomItem();// ÿһ��bomitem��Ϣ����
		Workbook workbook;
		Sheet sheet = null;
		 
		try {
			 
			is = new FileInputStream(excelFile);// ��ȡ�ļ�������
			// HSSFWorkbook workbook = new HSSFWorkbook(is);// ����Excel2003�ļ�����
			   workbook = new XSSFWorkbook(is);// ����Excel2003�ļ�����
			  sheet = workbook.getSheetAt(0);// ȡ����һ��������������0
		
		} catch (OLE2NotOfficeXmlFileException e) {
			 
			try {
				is = new FileInputStream(excelFile);// ��ȡ�ļ�������
				workbook = new HSSFWorkbook(is);
				sheet = workbook.getSheetAt(0);// ȡ����һ��������������0
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}// ����Excel2003�ļ�����
			   
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
		
		
		finally {// �ر��ļ���
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		
		
		
	
		// ��ʼѭ�������У���ͷ��������1��ʼ

		
		for (int i = 0; i <= sheet.getLastRowNum(); i++) {

			// HSSFRow row = sheet.getRow(i);// ��ȡ�ж���
			 Row row = sheet.getRow(i);// ��ȡ�ж���

			if (row == null) {// ���Ϊ�գ�������
				continue;
			}
			// ѭ��������Ԫ��
			BomItem bomitem1 = new BomItem();// ʵ����bomitem����
			for (int j = 0; j < row.getLastCellNum(); j++) {
				
				// HSSFCell cell = row.getCell(j);// ��ȡ��Ԫ�����
				 Cell cell = row.getCell(j);// ��ȡ��Ԫ�����

				if (cell == null) {// ��Ԫ��Ϊ������cellStrΪ�մ�
					cellStr = "";
				} else if (cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {// �Բ���ֵ�Ĵ���
					cellStr = String.valueOf(cell.getBooleanCellValue());
				} else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {// ������ֵ�Ĵ���
					cellStr =(int) cell.getNumericCellValue() + "";
				} else {// ���ఴ���ַ�������
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
		 
			bomItemList.add(bomitem1);// ����װ��List

		} // end of row
		
		
		
		System.out.print(bomItemList);
		 
	}

	 
}
