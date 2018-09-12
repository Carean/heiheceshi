package com.yansheng.core;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Map;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.yansheng.gui.GUI;
import com.yansheng.utils.filterByXlsx;

public class CreateWord {
	/**
	 * 
	 * @param destFile
	 * @param fileCon
	 */
	public boolean exportDoc(String destFile, String fileCon) {
		try {
//			 // doc content
//			 ByteArrayInputStream bais = new
//			 ByteArrayInputStream(fileCon.getBytes("gb2312"));
//			 POIFSFileSystem fs = new POIFSFileSystem();
//			 DirectoryEntry directory = fs.getRoot();
//			 directory.createDocument("WordDocument", bais);
//			 FileOutputStream ostream = new FileOutputStream(destFile);
//			 fs.writeFilesystem(ostream);
//			 // System.out.println(destFile+"导出成功！");
//			 bais.close();
//			 ostream.close();

			// Blank Document
			XWPFDocument document = new XWPFDocument();
			// Write the Document in file system
			FileOutputStream out = new FileOutputStream(destFile); // 下载路径/文件名称
			Calendar cal=Calendar.getInstance();    
		    int y=cal.get(Calendar.YEAR);    
		    int m=cal.get(Calendar.MONTH);    
		    int d=cal.get(Calendar.DATE);
		    //获取当前日期
		    String date=(m+1)+"月"+d+"日";
		    //截取文档标题
			String title=fileCon.substring(0, fileCon.indexOf(date)+4);
			//创建标题段落
			XWPFParagraph TitleParagraph = document.createParagraph();
			XWPFRun run = TitleParagraph.createRun();
			run.setText(title);
			run.setFontFamily("微软雅黑");
			run.setFontSize(14);
			//创建普通段落
			XWPFParagraph Paragraph = document.createParagraph();
			XWPFRun run2 = Paragraph.createRun();
			run2.setText(fileCon.substring(fileCon.indexOf(date)+4, fileCon.length()));
			//run2.addCarriageReturn();//增加回车换行
			run2.setFontFamily("微软雅黑");
			run2.setFontSize(10);
			document.write(out);
			out.close();
			return true;
		} catch (IOException e) {
			e.printStackTrace();
		}
		return false;
	}

	/**
	 * 去除文件后缀名
	 */
	public static String RemoveFileSuffix(String filename) {
		return filename.substring(0, filename.lastIndexOf("."));
	}
	
	/**
	 * 拼装日期
	 * @param month  8
	 * @param day    2
	 * @return       08.02
	 */
	public static String RefactorDate(String month,String day){
		String strDate="";
		if(Integer.parseInt(month)<10){
			strDate="0"+month;
		}else{
			strDate=month;
		}
		if(Integer.parseInt(day)<10){
			strDate=strDate+".0"+day;
		}else{
			strDate=strDate+"."+day;
		}
		return strDate;
	}

	public boolean CreateWord(String inputPath,String outputPath,String month,String day) {
		// 目标文件存放目录
		String destFile = outputPath;
		// 将要读取的excel目标的目录封装成 File对象。
		File dir = new File(inputPath);
		// 通过过滤器获取目录下的所有的 .xlsx 文件
		String[] pathArray = dir.list(new filterByXlsx());
		if(pathArray.length<=0){
			JOptionPane.showMessageDialog(new JFrame().getContentPane(), "当前目录下没有.xlsx文件", "错误",
					JOptionPane.ERROR_MESSAGE);
			return false;
		}
		// 读取表格类的对象
		ReadExcel excelData = null;
		// 存储排期个数
		int count = 0;
		// 存储每个点位的信息
		String statisticalData = "";
		// 统计当天总的点位个数
		int sum = 0;
		// 存储word文档导出情况
		String exportStatus = "";
		// 遍历指定目录下所有xlsx文件
		for (String path : pathArray) {
			excelData = new ReadExcel();
			// System.out.println("C:/Users/Administrator/Desktop/test/excel/"+path);
			// 读取表格获取表格数据
			List<List<String>> list = excelData.read(inputPath+"/"+ path);
			// 获取当天的点位数据
			ArrayList<Map<String, String>> maplist = excelData.getPointPositionValue(list,month, day);
			// ArrayList没数据说明当天没点位
			if (maplist.size() <= 0) {
				exportStatus += destFile + "\\" + path + "\t当天没有排期！\r\n\r\n";
				continue;
			} else {
				//排期数加1
				count++;
				String fileCon = "";
				// 单个点位数据的集合
				Map<String, String> map = null;
				// 遍历全部点位数据并逐一取出
				for (int index = 0; index < maplist.size(); index++) {
					map = maplist.get(index);
					if (index == 0) {
						fileCon += (path.substring(path.indexOf("年") - 4, path.indexOf("排期") + 2) + map.get("date")+"\r\r");
					}
					fileCon += (map.get("date"));
					fileCon += ("    ");
					fileCon += (map.get("mediaName"));
					fileCon += ("    ");
					fileCon += (map.get("terminal"));
					fileCon += ("    ");
					fileCon += (map.get("position"));
					fileCon += ("    ");
					fileCon += (map.get("form"));
					fileCon += ("\r\r\r\r");
					fileCon += ("落地页");
					fileCon += ("\r\r\r\r\r\r");
					// 控制其只执行一次
					if (index == 0) {
						statisticalData += path + "    点位数：" + map.get("count") + "\r\n\r\n";
					}
				}
				// 统计总点位数
				if (maplist.size() > 0) {
					sum += Integer.parseInt(map.get("count"));
				}
				// 创建word文档
				boolean exportDoc = new CreateWord().exportDoc(destFile + "/" + RemoveFileSuffix(path)
						+ RefactorDate(month, day) + ".doc", fileCon);
				if (exportDoc) {
					exportStatus += destFile + "\\" + path + "\t导出成功！\r\n\r\n";
					GUI.console.append(exportStatus);
					GUI.console.paintImmediately(GUI.console.getBounds());//立即刷新文本域的值
					GUI.consolePanel.paintImmediately(GUI.consolePanel.getBounds());
					System.out.println("success");
				} else {
					exportStatus += destFile + "\\" + path + "\t导出失败！\r\n\r\n";
					GUI.console.append(exportStatus);
					GUI.console.paintImmediately(GUI.console.getBounds());//立即刷新文本域的值
					GUI.consolePanel.paintImmediately(GUI.consolePanel.getBounds());
					System.out.println("failure");
				}
			}
		}
		// 加上总点位
		statisticalData += "当前日期：" + new SimpleDateFormat("yyyy/MM/dd").format(new Date()) + "\t" + month + "月" + day
				+ "日排期个数：" + count + "\t总点位数：" + sum;
		// 写入点位数据
		try {
			FileOutputStream outStr = new FileOutputStream(new File(destFile + "/word点位统计"+ month + "月" + day+"日.txt"));
			BufferedOutputStream Buff = new BufferedOutputStream(outStr);
			Buff.write(statisticalData.getBytes());
			Buff.flush();
			Buff.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		// 写入文件导出情况
		try {
			FileOutputStream outSTr = new FileOutputStream(new File(destFile + "/word文档导出情况"+ month + "月" + day+"日.txt"));
			BufferedOutputStream Buff = new BufferedOutputStream(outSTr);
			Buff.write(exportStatus.getBytes());
			Buff.flush();
			Buff.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return true;
	}

}
