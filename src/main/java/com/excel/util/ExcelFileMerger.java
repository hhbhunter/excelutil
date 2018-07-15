package com.excel.util;
/**
 * @author hhbhunter
 * 2018-01-09
 */
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileMerger {
	 static SimpleDateFormat  sdf = new SimpleDateFormat("yyyy/MM/dd");
	public static XSSFWorkbook readexcel(String fileName) throws InvalidFormatException, IOException{
		File file = new File(fileName);
	    return new XSSFWorkbook(file);
	}
	public static void closeExecle(XSSFWorkbook excel) throws IOException{
		excel.close();
	}
	
		
	public static List<XSSFRow> readExecleSheet(XSSFWorkbook excel,String sheetName,int startNum) throws InvalidFormatException, IOException {
		
	    XSSFSheet xssfSheet = excel.getSheet(sheetName);
	    if(xssfSheet==null) {
	    	return null;
	    }
	    
	    int rowstart = xssfSheet.getFirstRowNum();
	    if(startNum>0) rowstart=startNum;
	    int rowEnd = xssfSheet.getLastRowNum();
	    StringBuffer allStr = new StringBuffer();
	    List<XSSFRow> sheetData=new ArrayList<XSSFRow>();
	    for(int i=rowstart;i<=rowEnd;i++)
	    {
	        XSSFRow row = xssfSheet.getRow(i);
	        if(null == row) continue;
	        sheetData.add(row);
	        StringBuffer stringBuffer = new StringBuffer();
			stringBuffer.append("\n");
	        allStr.append(stringBuffer);
	    }
	   return sheetData;
	  	        	
	}
	
	
	public static void writeToXlsx(String readsheetName,String workBookName,int startNum) throws Exception{
		String outFile=System.getProperty("user.dir")+File.separator+workBookName;
		int lineNum=0;
		File outXss=new File(outFile);
		OutputStream os = null;
		//创建excel文件  
		 XSSFWorkbook xssf_w_book;
		 XSSFSheet xssfSheet = null;
		if(outXss.isFile() && outXss.exists()){
			System.out.println(outFile+" 存在 。。。将追加内容");
			//创建excel文件  
			FileInputStream fis = new FileInputStream(outFile);
	        xssf_w_book=new XSSFWorkbook(fis); 
	        xssfSheet = xssf_w_book.getSheet(readsheetName);
	        if(xssfSheet!=null)
	        lineNum= xssfSheet.getLastRowNum()+1;
	        System.out.println(readsheetName+" 当前有 "+lineNum+" 行");
		}else{
			xssf_w_book=new XSSFWorkbook();
		}
		//输出流定义  
		os = new FileOutputStream(outFile);
		String dataDir=System.getProperty("user.dir")+File.separator+"data";
        
        File dir=new File(dataDir);
        String[] files={};
        if(dir.isDirectory() ){
        	files=dir.list(new FilenameFilter() {
				
				public boolean accept(File dir, String name) {
					// TODO Auto-generated method stub
					if(name.lastIndexOf("xlsx")>0)return true;
					return false;
				}
			});
        }else{
        	System.out.println("不存在目录:"+System.getProperty("user.dir")+File.separator+"data  请创建！");
        }
        
        System.out.println("总共存在.xlsx文件个数"+files.length);
        for(String file:files){
        	XSSFWorkbook otherBook=readexcel(dataDir+File.separator+file);
        	List<XSSFRow> sheetData=readExecleSheet(otherBook, readsheetName, startNum);
        	if(sheetData==null){
        		System.out.println(file+" 没有 "+readsheetName);
        	}
        	int current=0;
        	current=writeToSheet(xssf_w_book, readsheetName, sheetData, lineNum,file);
        	System.out.println(file+" 总行数："+(current-lineNum));
        	closeExecle(otherBook);
        	lineNum=current;
        }
      //excel文件导出  
        xssf_w_book.write(os);  
        os.flush();
        os.close();  
        System.out.println(workBookName+"文件追加完成,总行数"+lineNum);
	}
	private static int writeToSheet(XSSFWorkbook xssf_w_book,String sheetName,List<XSSFRow> sheetData,int lineNum,String file){
		XSSFSheet xssf_w_sheet=xssf_w_book.getSheet(sheetName);
		if(xssf_w_book.getSheet(sheetName)==null){
			
			xssf_w_sheet=xssf_w_book.createSheet(sheetName);
		}
        xssf_w_sheet.setDefaultColumnWidth(21); //固定列宽度  
        XSSFRow xssf_w_row=null;//创建一行  
        XSSFCell xssf_w_cell=null;//创建每个单元格  
        //定义表头单元格样式  
        XSSFCellStyle head_cellStyle=xssf_w_book.createCellStyle();  
        //定义表头字体样式  
        XSSFFont  head_font=xssf_w_book.createFont();  
        head_font.setFontName("宋体");//设置头部字体为宋体  
        head_font.setBoldweight(Font.BOLDWEIGHT_BOLD); //粗体  
        head_font.setFontHeightInPoints((short) 10);  
        //表头单元格样式设置  
        head_cellStyle.setFont(head_font);//单元格使用表头字体样式  
        head_cellStyle.setAlignment(HorizontalAlignment.CENTER);  
        head_cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);  
        head_cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
        head_cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);  
        head_cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);  
        head_cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);  
        if(sheetData==null)return lineNum;
        for(XSSFRow rowData:sheetData){
        	int cellEnd = rowData.getLastCellNum();
        	if(cellEnd!=0){
        		if(rowData.getCell(0)==null){
        			continue;
        		}else{
        			if(rowData.getCell(0).getStringCellValue().equalsIgnoreCase("-")||rowData.getCell(0).getStringCellValue().equalsIgnoreCase(""))continue;
        		}
        	}else{
        		continue;
        	}
        	xssf_w_row=xssf_w_sheet.createRow(lineNum);
        	for(int i=0;i<cellEnd;i++){
        		 xssf_w_cell = xssf_w_row.createCell(i);
        		 if(rowData.getCell(i)==null)continue;
        		try{
        			 xssf_w_cell.setCellValue(rowData.getCell(i).getStringCellValue());
//        			 System.out.println(rowData.getCell(i).getStringCellValue());
        		 }catch(Exception e){
        			 try{
        			 if (HSSFDateUtil.isCellDateFormatted(rowData.getCell(i))) {
//							System.out.println(sdf.format(rowData.getCell(i).getDateCellValue()));
	                		xssf_w_cell.setCellValue(sdf.format(rowData.getCell(i).getDateCellValue()));
						}else {
							xssf_w_cell.setCellValue(rowData.getCell(i).getNumericCellValue()); 
//							System.out.println(rowData.getCell(i).getNumericCellValue());
						}
        			 }catch(Exception e1){
        				 System.out.println("================================");
        				 System.out.println("【错误信息】["+file+"]的【"+sheetName+"】第"+(rowData.getRowNum()+1)+"行的"+(i+1)+"单元格有问题");
        				 System.out.println("================================");
        			 }
        		 }
                  
        	}
        	lineNum++;
        }
        
		return lineNum;
	}
	
	
	
	
	public static void main(String[] args) throws Exception {
		String filename="test.xlsx";
		String readsheetName="";
		int startLineNum=2;
		if(args.length<3){
			System.out.println("usage: input [filename],[readsheetName],[startLineNum]");
			System.exit(0);
		}else{
			startLineNum=Integer.parseInt(args[2]);
			filename=args[0];
			readsheetName=args[1];
		}
		writeToXlsx(readsheetName, filename, startLineNum);
	}
	
}
