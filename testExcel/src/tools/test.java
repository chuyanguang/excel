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
		//����HSSFWorkbook����(excel���ĵ�����)  
				HSSFWorkbook wkb = new HSSFWorkbook();  
				//�����µ�sheet����excel�ı���  
				HSSFSheet sheet=wkb.createSheet("�ɼ���");  
				//��sheet�ﴴ����һ�У�����Ϊ������(excel����)��������0��65535֮����κ�һ��  
				HSSFRow row1=sheet.createRow(0);  
				//������Ԫ��excel�ĵ�Ԫ�񣬲���Ϊ��������������0��255֮����κ�һ��  
				HSSFCell cell=row1.createCell(0);  
				      //���õ�Ԫ������  
				cell.setCellValue("ѧԱ���Գɼ�һ����");  
				//�ϲ���Ԫ��CellRangeAddress����������α�ʾ��ʼ�У������У���ʼ�У� ������  
				sheet.addMergedRegion(new CellRangeAddress(0,0,0,3));  
				//��sheet�ﴴ���ڶ���  
				HSSFRow row2=sheet.createRow(1);      
				      //������Ԫ�����õ�Ԫ������  
				      row2.createCell(0).setCellValue("����");  
				      row2.createCell(1).setCellValue("�༶");      
				      row2.createCell(2).setCellValue("���Գɼ�");  
				row2.createCell(3).setCellValue("���Գɼ�");      
				      //��sheet�ﴴ��������  
				      HSSFRow row3=sheet.createRow(2);  
				      row3.createCell(0).setCellValue("����");  
				      row3.createCell(1).setCellValue("As178");  
				      row3.createCell(2).setCellValue(87);      
				      row3.createCell(3).setCellValue(78);      
				  //.....ʡ�Բ��ִ���  
				  
				  
				//���Excel�ļ�  
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
				    	  System.out.println("д��ɹ���");
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					} 
	}

	public static List<ScoreInfo> loadScoreInfo(String xlsPath) throws IOException{  
		List temp = new ArrayList<>();
		FileInputStream fileIn = new FileInputStream(xlsPath);
		// ����ָ�����ļ�����������Excel�Ӷ�����Workbook����
		Workbook wb0 = new HSSFWorkbook(fileIn);
		// ��ȡExcel�ĵ��еĵ�һ����
		Sheet sht0 = wb0.getSheetAt(0);
		// ��Sheet�е�ÿһ�н��е���
		for (Row r : sht0) {
			// �����ǰ�е��кţ���0��ʼ��δ�ﵽ2�������У������ѭ��
			if (r.getRowNum() < 2) {
				continue;
			}
			// ����ʵ����
			ScoreInfo info = new ScoreInfo();
			// ȡ����ǰ�е�1����Ԫ�����ݣ�����װ��infoʵ��stuName������
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
