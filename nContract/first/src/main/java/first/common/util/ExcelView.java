 package first.common.util;

import java.net.URLEncoder;
import java.util.*;
import java.util.Iterator;
import java.util.Map.*;
import java.util.Set;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
//import javax.swing.text.html.HTMLDocument.Iterator;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;
import org.springframework.web.servlet.view.document.AbstractExcelView;

@Component("excelView")
public class ExcelView extends AbstractExcelView {

	//declare global variables.
	String excelName = null;
	HSSFSheet worksheet = null;
    HSSFRow row = null;
    TreeMap<String,Object> treeMap = null;
    
    
	
	//@SuppressWarnings("unchecked")
	@Override
	protected void buildExcelDocument(Map<String, Object> model
									  , HSSFWorkbook workbook
									  , HttpServletRequest req,
									  HttpServletResponse res) throws Exception {
		
		switch (model.keySet().toString().toUpperCase())
		{
			case "[LIST]" : 
				setContractList(model, workbook, req, res);
				break;
		
			case "[ADMINLIST]" :
				setContractAdminCodeList(model, workbook, req, res);
				break;
		}
        
        res.setContentType("Application/Msexcel");
        res.setHeader("Content-Disposition", "ATTachment; Filename="+excelName);
  }
	
	//this is data of all Contract list(All user can use.) 
	public void setContractList(Map<String, Object> model
								, HSSFWorkbook workbook
								, HttpServletRequest req
								, HttpServletResponse res) throws Exception{
		
		excelName=URLEncoder.encode("Contract_All_List.xls","UTF-8");										//file name of Excle
		worksheet = workbook.createSheet(excelName.substring(0, excelName.indexOf("."))+ " WorkSheet");		//Sheet name of Excel
		
		@SuppressWarnings("unchecked")																		
		List<Map<String, Object>> list = (List<Map<String, Object>>)model.get("list");						//store in list collection the received delivery information. 
		
		//set first column name of Excel.
        row = worksheet.createRow(0);
        row.createCell(0).setCellValue("사업부");
        row.createCell(1).setCellValue("계약번호");
        row.createCell(2).setCellValue("구분");
        row.createCell(3).setCellValue("고객사_1");
        row.createCell(4).setCellValue("고객사_2");
        row.createCell(5).setCellValue("계약구분");
        row.createCell(6).setCellValue("계약서명");
        row.createCell(7).setCellValue("목적사업");
        row.createCell(8).setCellValue("효력발생일");
        row.createCell(9).setCellValue("기간만료일");
        row.createCell(10).setCellValue("해지조건");
        row.createCell(11).setCellValue("자동연장기간");
        row.createCell(12).setCellValue("해지통지기간");
        row.createCell(13).setCellValue("해지여부");
        row.createCell(14).setCellValue("부속계약서");
        row.createCell(15).setCellValue("비고");
		
        setExcelBinding(list);
	}

	//this is data of various standard information list(administrator user can use.) 
	public void setContractAdminCodeList(Map<String, Object> model, HSSFWorkbook workbook, HttpServletRequest req,
			HttpServletResponse res) throws Exception{

		excelName=URLEncoder.encode("Contract_AdminCode_List.xls","UTF-8");
		worksheet = workbook.createSheet(excelName.substring(0, excelName.indexOf("."))+ " WorkSheet");

		@SuppressWarnings("unchecked")
		List<Map<String, Object>> list = (List<Map<String, Object>>)model.get("adminList");

		//set first column name of Excel.
		row = worksheet.createRow(0);
		row.createCell(0).setCellValue("기준정보 구분");
        row.createCell(1).setCellValue("Code");
        row.createCell(2).setCellValue("Code 명");
        row.createCell(3).setCellValue("약자");
        row.createCell(4).setCellValue("상위 Code");
        row.createCell(5).setCellValue("Level");
        
        setExcelBinding(list);
	}
	
	public void setExcelBinding(List<Map<String, Object>> excelList){
		int rowCnt = 1;
		
        for (Map<String, Object> RowMap : excelList) {
    		int colCnt = 0;
    		
        	row = worksheet.createRow(rowCnt); 						//create new excel row (data row after title row)
        	treeMap = new TreeMap<String,Object>(RowMap); 			//create an object for sorting
 
        	Set<Entry<String, Object>> set = treeMap.entrySet();	//create an 'Set' Object for to use 'iterator'. 
        	Iterator<Entry<String, Object>> itr = set.iterator();	
        
        	//repeat until there is no next iterator data.
        	while (itr.hasNext())
        	{
        		Map.Entry<String, Object> e = (Map.Entry<String, Object>)itr.next();
        		row.createCell(colCnt).setCellValue(e.getValue().toString());
        		
        		colCnt++;
	    	}
        	
        	rowCnt++;
		}
	}
	
}
