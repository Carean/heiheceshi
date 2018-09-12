package com.yansheng.core;

import java.awt.geom.Rectangle2D;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
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

import org.apache.poi.common.usermodel.fonts.FontGroup;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

import com.yansheng.gui.GUI;
import com.yansheng.utils.filterByXlsx;

public class CreatePPT {
	/**
	 * 创建第一张幻灯片（展示排期信息）
	 * 
	 * @param ppt
	 * @param headText
	 */
	public void CreateFirstSlide(XMLSlideShow ppt, String headText,String pngPath,String month,String day) {
		try {
			// 创建一张幻灯片
			XSLFSlide slide = ppt.createSlide();
			// 转义图片路径
			String path = pngPath;
			String newPath = path.replace("\\\\", "/");
			// 封装图片为File对象
			File image = new File(newPath);
			// 读取图片
			byte[] picture = IOUtils.toByteArray(new FileInputStream(image));
			// 将图片添加到ppt中
			XSLFPictureData idx = ppt.addPicture(picture, PictureData.PictureType.PNG);
			// 在PPT中创建图片
			XSLFPictureShape pic = slide.createPicture(idx);
			// 创建展示标题的文本框
			XSLFTextBox textBox = slide.createTextBox();
			// 设置文本框位置及大小
			textBox.setAnchor(new Rectangle2D.Double(0, 180, 730, 140));
			// 创建段落
			XSLFTextParagraph Paragraph = textBox.addNewTextParagraph();
			XSLFTextRun Run = Paragraph.addNewTextRun();
			// 设置文本内容
			Run.setText(headText);
			// 设置字体
			Run.setFontFamily("微软雅黑", FontGroup.LATIN);
			// 设置字体加粗
			Run.setFontColor(java.awt.Color.CYAN);
			// 设置字体颜色
			Run.setFontColor(java.awt.Color.white);
			// 设置字体大小
			Run.setFontSize(40.00);
			// 创建展示日期的文本框
			XSLFTextBox textBox2 = slide.createTextBox();
			// 设置文本框位置及大小
			textBox2.setAnchor(new Rectangle2D.Double(540, 370, 140, 40));
			// 创建段落
			XSLFTextParagraph Paragraph2 = textBox2.addNewTextParagraph();
			XSLFTextRun Run2 = Paragraph2.addNewTextRun();
			Calendar cal = Calendar.getInstance();
			int y = cal.get(Calendar.YEAR);
			// 设置文本内容
			Run2.setText(y + "年" + month + "月" + day + "日");
			// 设置字体
			Run2.setFontFamily("微软雅黑", FontGroup.LATIN);
			// 设置字体加粗
			Run2.setFontColor(java.awt.Color.CYAN);
			// 设置字体颜色
			Run2.setFontColor(java.awt.Color.black);
			// 设置字体大小
			Run2.setFontSize(16.00);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * 创建其他普通幻灯片
	 * 
	 * @param ppt
	 * @param text
	 */
	public void CreateOtherSlide(XMLSlideShow ppt, String text,String pngPath) {
		try {
			// 创建一张幻灯片
			XSLFSlide slide = ppt.createSlide();
			// 转义图片路径
			String path = pngPath;
			String newPath = path.replace("\\\\", "/");
			// 封装图片为File对象
			File image = new File(newPath);
			// 读取图片
			byte[] picture = IOUtils.toByteArray(new FileInputStream(image));
			// 将图片添加到ppt中
			XSLFPictureData idx = ppt.addPicture(picture, PictureData.PictureType.PNG);
			// 在PPT中创建图片
			XSLFPictureShape pic = slide.createPicture(idx);
			// 创建文本框
			XSLFTextBox textBox = slide.createTextBox();
			// 设置文本框位置及大小
			textBox.setAnchor(new Rectangle2D.Double(14, 46, 730, 30));
			// 创建段落
			XSLFTextParagraph Paragraph = textBox.addNewTextParagraph();
			XSLFTextRun Run = Paragraph.addNewTextRun();
			// 设置文本内容
			Run.setText(text);
			// 设置字体
			Run.setFontFamily("微软雅黑", FontGroup.LATIN);
			// 设置字体大小
			Run.setFontSize(14.00);
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	/**
	 * 创建公共的幻灯片（落地页）
	 * 
	 * @param ppt
	 */
	public void CreateCommonSlide(XMLSlideShow ppt,String pngPath) {
		try {
			// 创建一张幻灯片
			XSLFSlide slide = ppt.createSlide();
			// 转义图片路径
			String path = pngPath;
			String newPath = path.replace("\\\\", "/");
			// 封装图片为File对象
			File image = new File(newPath);
			// 读取图片
			byte[] picture = IOUtils.toByteArray(new FileInputStream(image));
			// 将图片添加到ppt中
			XSLFPictureData idx = ppt.addPicture(picture, PictureData.PictureType.PNG);
			// 在PPT中创建图片
			XSLFPictureShape pic = slide.createPicture(idx);
			// 创建文本框
			XSLFTextBox textBox = slide.createTextBox();
			// 设置文本框位置及大小
			textBox.setAnchor(new Rectangle2D.Double(14, 46, 730, 30));
			// 创建段落
			XSLFTextParagraph Paragraph = textBox.addNewTextParagraph();
			XSLFTextRun Run = Paragraph.addNewTextRun();
			// 设置文本内容
			Run.setText("落地页");
			// 设置字体
			Run.setFontFamily("微软雅黑", FontGroup.LATIN);
			// 设置字体大小
			Run.setFontSize(14.00);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * 生成ppt文件
	 * 
	 * @param ppt
	 * @param destFile
	 * @return
	 */
	public boolean CreatePPTFile(XMLSlideShow ppt, String destFile) {
		try {
			FileOutputStream out = new FileOutputStream(new File(destFile));
			ppt.write(out);
			// System.out.println("Presentation created successfully");
			out.close();
			return true;
		} catch (IOException e) {
			e.printStackTrace();
		}
		return false;
	}

	public boolean CreatePPT(String inputPath, String outputPath,String firstpngPath,String otherpngPath, String month, String day) {
		// 目标文件存放目录
		String destFile = outputPath;
		// 将要读取的excel目标的目录封装成 File对象。
		File dir = new File(inputPath);
		// 通过过滤器获取目录下的所有的 .xlsx 文件
		String[] pathArray = dir.list(new filterByXlsx());
		if (pathArray.length <= 0) {
			JOptionPane.showMessageDialog(new JFrame().getContentPane(), "当前目录下没有.xlsx文件", "错误",
					JOptionPane.ERROR_MESSAGE);
			return false;
		}
		// 创建读取表格类的对象
		ReadExcel excelData = null;
		// 创建ppt文档
		CreatePPT createPPT = new CreatePPT();
		// 创建一个空的ppt文件对象
		XMLSlideShow ppt = null;
		// 存储排期个数
		int count = 0;
		// 存储每个点位的信息
		String statisticalData = "";
		// 统计点位当天点位个数
		int sum = 0;
		// 遍历指定目录下所有xlsx文件
		// 存储ppt文件导出情况
		String exportStatus = "";
		for (String path : pathArray) {
			// 初始化ppt对象
			ppt = new XMLSlideShow();
			// 初始化读取表格类的对象
			excelData = new ReadExcel();
			// 读取表格获取表格数据
			List<List<String>> list = excelData.read(inputPath + "/" + path);
			// 获取某月某日值为1的点位数据
			ArrayList<Map<String, String>> maplist = excelData.getPointPositionValue(list, month, day);
			// ArrayList没数据说明当天没点位
			if (maplist.size() <= 0) {
				exportStatus += destFile + "/" + path + "\t当天没有排期！\r\n\r\n";
				continue;
			} else {
				count++;
				// 单个点位数据
				String pointLocationData = "";
				// 单个点位数据集合
				Map<String, String> map = null;
				// 创建第一张幻灯片（展示排期信息）
				createPPT.CreateFirstSlide(ppt, (path.substring(path.indexOf("年") - 4, path.indexOf("排期") + 2)),firstpngPath,month,day);
				// 遍历全部点位数据并逐一取出并创建幻灯片
				for (int index = 0; index < maplist.size(); index++) {
					map = maplist.get(index);
					pointLocationData += (map.get("date"));
					pointLocationData += ("\t");
					pointLocationData += (map.get("mediaName"));
					pointLocationData += ("\t");
					pointLocationData += (map.get("terminal"));
					pointLocationData += ("\t");
					pointLocationData += (map.get("position"));
					pointLocationData += ("\t");
					pointLocationData += (map.get("form"));
					// 创建一张幻灯片展示点位
					createPPT.CreateOtherSlide(ppt, pointLocationData,otherpngPath);
					// 创建一张幻灯片展示落地页
					createPPT.CreateCommonSlide(ppt,otherpngPath);
					// 控制其只执行一次
					if (index == 0) {
						statisticalData += path + "    点位数：" + map.get("count") + "\r\n\r\n";
					}
					// 重置点位数据
					pointLocationData = "";
				}
				// 统计点位数
				if (maplist.size() > 0) {
					sum += Integer.parseInt(map.get("count"));
				}

				// 生成PPT文件
				boolean createPPTFile = createPPT.CreatePPTFile(ppt, destFile + "/" + CreateWord.RemoveFileSuffix(path)
						+ CreateWord.RefactorDate(month, day) + ".pptx");

				if (createPPTFile) {
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
			FileOutputStream outSTr = new FileOutputStream(new File(destFile + "/ppt点位统计"+ month + "月" + day+"日.txt"));
			BufferedOutputStream Buff = new BufferedOutputStream(outSTr);
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
			FileOutputStream outSTr = new FileOutputStream(new File(destFile + "/ppt文件导出情况"+ month + "月" + day+"日.txt"));
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
