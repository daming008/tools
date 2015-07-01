package com.dm.readfile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import org.junit.Test;
/**
 * 
 * @author ldm 973264333@qq.com 2015.07.01
 *用到的jar包jexcelapi_2_6_6.zip在360网盘“所有文件\myWork\”目录中
 */

public class ExcelRead {
	@Test
	public void read(){
		Workbook readwb = null;
		try {
			InputStream instream = new FileInputStream("e:/excelFile/注册数据.xls");
			readwb = Workbook.getWorkbook(instream);
			
			//获取第一张sheeet表
			Sheet readSheet = readwb.getSheet(0);
			//获取总列数
			int rsColumns = readSheet.getColumns();
			//获取总行数
			int rsRows = readSheet.getRows();
			
			for(int i=0;i<rsRows;i++){
				for(int j=0;j<rsColumns;j++){
					Cell cell = readSheet.getCell(j, i);
					System.out.print(cell.getContents()+" :");
				}
				System.out.println();
			}
	
			WritableWorkbook wwb = Workbook.createWorkbook(new File("e:/excelFile/红楼梦1.xls"), readwb);
			WritableSheet ws= wwb.getSheet(0);
			WritableCell wc = ws.getWritableCell(1, 0);
			if(wc.getType() == CellType.LABEL){
				Label l = (Label)wc;
				l.setString("新姓名");
			}
			wwb.write();
			wwb.close();
		} catch (Exception e) {
			e.printStackTrace();
		}finally{
			readwb.close();
		}
	}
	
	@Test
	public void testWrite(){
		WritableWorkbook book = null;
		try {
			book = Workbook.createWorkbook(new File("e:/excelFile/测试.xls"));
			WritableSheet sheet = book.createSheet("第一页",0);
			Label label = new Label(0, 0, "测试");
			
			WritableFont wfc = new WritableFont(WritableFont.ARIAL,10,WritableFont.NO_BOLD,false,UnderlineStyle.NO_UNDERLINE,jxl.format.Colour.DARK_YELLOW); 
			WritableCellFormat wcfFC = new WritableCellFormat(wfc);
			sheet.addCell(label);
			
			Number number = new Number(1, 0, 123.456,wcfFC);
			sheet.addCell(number);
			
			Label s = new Label(1, 2, "三十三");
			sheet.addCell(s);
			book.write();
		} catch (Exception e) {
			e.printStackTrace();
		}finally{
			try {
				book.close();
			} catch (WriteException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
	@Test
	public void testWritePaperless(){
		String[] options = {"A","B","C","D","E"};
		String[] difficul ={"0","0.1","0.2","0.3"};
		WritableWorkbook book = null;
		int count = 10;
		int countKonwledge = 3;
		
		try {
			WritableFont fontTitle = new WritableFont(WritableFont.ARIAL,9,WritableFont.BOLD,false,UnderlineStyle.NO_UNDERLINE,jxl.format.Colour.BLACK);  
			fontTitle.setColour(jxl.format.Colour.RED);  
			WritableCellFormat formatTitle = new WritableCellFormat(fontTitle);  
			formatTitle.setAlignment(Alignment.CENTRE);
			
//			CellView navCellView = new CellView();  
//		    navCellView.setAutosize(true); //设置自动大小
//		    navCellView.setSize(28);
			
			book = Workbook.createWorkbook(new File("e:/excelFile/测试.xls"));
			WritableSheet sheet = book.createSheet("单选题", 0);
			sheet.setColumnView(1, 30);//设置题干这一列的单元格的宽度
			sheet.setColumnView(4, 25);//设置答案这一列的单元格的宽度
			
			Label number = new Label(0, 0, "序号",formatTitle);
			Label context = new Label(1, 0, "题干",formatTitle);
			Label knowledge = new Label(2, 0, "知识点",formatTitle);
			Label option = new Label(3, 0, "选项",formatTitle);
			Label answer = new Label(4, 0, "答案",formatTitle);
			Label difficulty = new Label(5, 0, "难度",formatTitle);
			Label distinciton = new Label(6, 0, "区分度",formatTitle);
			
			
			sheet.addCell(number);
			sheet.addCell(context);
			sheet.addCell(knowledge);
			sheet.addCell(option);
			sheet.addCell(answer);
			sheet.addCell(difficulty);
			sheet.addCell(distinciton);
			
			for(int d=0;d<difficul.length*count*countKonwledge;d++){
				Label diffic = new Label(5, (options.length*d+1), difficul[d%difficul.length]);
				sheet.addCell(diffic);
				for(int j=0;j<countKonwledge*count*difficul.length;j++){
					Label konw = new Label(2, options.length*j+1, "知识点"+(j/count+1/difficul.length)+"");
					sheet.addCell(konw);
					for(int i=1;i<((count+1)*countKonwledge-2)*difficul.length;i++){
						Number serial = new Number(0, 1+(i-1)*options.length, i);
						Label cont = new Label(1, 1+(i-1)*options.length,"我是单选题的第"+i+"题，知识点"+(j/count+1)+"");
						sheet.addCell(serial);
						sheet.addCell(cont);
					}
				}
			}
			
				
				for(int k=0;k<(count)*options.length*countKonwledge*difficul.length;k++){
					Label opt = new Label(3, k+1, options[k%options.length]);
					if(k%options.length == 2){
						Label ans = new Label(4, k+1, "答案是我是我就是我！！！");
						sheet.addCell(ans);
					}else{
						Label ans = new Label(4, k+1, "答案不是我不是我");
						sheet.addCell(ans);
					}
					sheet.addCell(opt);
				}
			
			
			book.write();
		} catch (Exception e) {
			e.printStackTrace();
		}finally{
			try {
				book.close();
			} catch (WriteException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
}
