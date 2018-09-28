package tools;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;


public class test {
	public static void main(String[] args) {
		
		try {
			List<ScoreInfo> list = loadScoreInfo("d:\\\\workbook.xls");
			System.out.println(list.size());
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public static void createExcel() {
		//创建HSSFWorkbook对象(excel的文档对象)  
				HSSFWorkbook wkb = new HSSFWorkbook();  
				//建立新的sheet对象（excel的表单）  
				HSSFSheet sheet=wkb.createSheet("成绩表");  
				//在sheet里创建第一行，参数为行索引(excel的行)，可以是0～65535之间的任何一个  
				HSSFRow row1=sheet.createRow(0);  
				//创建单元格（excel的单元格，参数为列索引，可以是0～255之间的任何一个  
				HSSFCell cell=row1.createCell(0);  
				      //设置单元格内容  
				cell.setCellValue("学员考试成绩一览表");  
				//合并单元格CellRangeAddress构造参数依次表示起始行，截至行，起始列， 截至列  
				sheet.addMergedRegion(new CellRangeAddress(0,0,0,3));  
				//在sheet里创建第二行  
				HSSFRow row2=sheet.createRow(1);      
				      //创建单元格并设置单元格内容  
				      row2.createCell(0).setCellValue("姓名");  
				      row2.createCell(1).setCellValue("班级");      
				      row2.createCell(2).setCellValue("笔试成绩");  
				row2.createCell(3).setCellValue("机试成绩");      
				      //在sheet里创建第三行  
				      HSSFRow row3=sheet.createRow(2);  
				      row3.createCell(0).setCellValue("李明");  
				      row3.createCell(1).setCellValue("As178");  
				      row3.createCell(2).setCellValue(87);      
				      row3.createCell(3).setCellValue(78);      
				  //.....省略部分代码  
				  
				  
				//输出Excel文件  
//				    OutputStream output=response.getOutputStream();  
//				    response.reset();  
//				    response.setHeader("Content-disposition", "attachment; filename=details.xls");  
//				    response.setContentType("application/msexcel");          
//				    wkb.write(output);  
//				    output.close();  
//				retrun null;
				      try {
				    	  FileOutputStream output=new FileOutputStream("d:\\workbook.xls");  
				    	  wkb.write(output);  
				    	  output.flush();
				    	  System.out.println("写入成功！");
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					} 
	}

	public static List<ScoreInfo> loadScoreInfo(String xlsPath) throws IOException{  
		List temp = new ArrayList<>();
		FileInputStream fileIn = new FileInputStream(xlsPath);
		// 根据指定的文件输入流导入Excel从而产生Workbook对象
		Workbook wb0 = new HSSFWorkbook(fileIn);
		// 获取Excel文档中的第一个表单
		Sheet sht0 = wb0.getSheetAt(0);
		// 对Sheet中的每一行进行迭代
		for (Row r : sht0) {
			// 如果当前行的行号（从0开始）未达到2（第三行）则从新循环
			if (r.getRowNum() < 2) {
				continue;
			}
			// 创建实体类
			ScoreInfo info = new ScoreInfo();
			// 取出当前行第1个单元格数据，并封装在info实体stuName属性上
			info.setStuName(r.getCell(0).getStringCellValue());
			info.setClassName(r.getCell(1).getStringCellValue());
			info.setRscore(r.getCell(2).getNumericCellValue());
			info.setLscore(r.getCell(3).getNumericCellValue());
			temp.add(info);
		}
		fileIn.close();
		return temp;
	}

}
