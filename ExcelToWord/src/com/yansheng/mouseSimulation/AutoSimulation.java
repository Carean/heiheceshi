package com.yansheng.mouseSimulation;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.InputEvent;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import com.yansheng.core.ReadExcel;
import com.yansheng.utils.City_Utils;
import com.yansheng.utils.CreateScreenCapture;
import com.yansheng.utils.OpenIE;
import com.yansheng.utils.filterByXlsx;

public class AutoSimulation {
	private int count = 0;

	/**
	 * 自动模拟拷屏
	 * 
	 * @param ExcelPath
	 *            Excel文件夹路径
	 */
	public void startSimulation(String ExcelPath, String pngPath,String batPath, String month, String day) {
		File file = new File(ExcelPath);
		if (file.isDirectory()) {
			// 通过过滤器获取目录下的所有的 .xlsx 文件
			String[] pathArray = file.list(new filterByXlsx());
			if (pathArray.length <= 0) {
				JOptionPane.showMessageDialog(new JFrame().getContentPane(), "当前目录下没有.xlsx文件", "错误",
						JOptionPane.ERROR_MESSAGE);
			} else {
				// 遍历指定目录下所有xlsx文件
				for (String path : pathArray) {
					if (path.indexOf("易车") >= 0) {
						ReadExcel excelData = new ReadExcel();
						// System.out.println("C:/Users/Administrator/Desktop/test/excel/"+path);
						// 读取表格获取表格数据
						List<List<String>> list = excelData.read(ExcelPath + "/" + path);
						// 获取当天的点位数据
						ArrayList<Map<String, String>> maplist = excelData.getPointPositionValue(list, month, day);
						// ArrayList没数据说明当天没点位
						if (maplist.size() <= 0) {
							continue;
						} else {
							// 遍历全部点位数据并逐一取出
							for (int index = 0; index < maplist.size(); index++) {
								Map<String, String> map = maplist.get(index);
								if (map.get("mediaName").indexOf("易车") >= 0) {
									if ("pc端".equals(map.get("terminal")) || "PC端".equals(map.get("terminal"))) {
										String form = map.get("form");
										// 点位名
										String ADName = form.substring(0, form.indexOf("/"));
										// 链接
										String url = map.get("url");
										// 点位城市拼音
										String cityName = new City_Utils().getCity(form);
										// 截图名
										String pngName = "";
										if (path.indexOf("、") >= 0) {
											pngName = path.substring(0, path.indexOf("、")) + "-";
										} else if (path.indexOf(".") >= 0) {
											pngName = path.substring(0, path.indexOf(".")) + "-";
										}
										// 自动截屏
										doPrintScreen(ADName, url, cityName, pngName, pngPath,batPath);
									}

								}
							}
						}

					}
					count = 0;
				}
			}
		}
	}

	/**
	 * 
	 * @param ADName
	 * @param url
	 * @param cityName
	 * @param pngName
	 * @param pngPath
	 * @param batPath
	 */
	public void doPrintScreen(String ADName, String url, String cityName, String pngName, String pngPath,String batPath) {
		File file = new File("C:\\Users\\Administrator\\Desktop\\config_yiche.txt");
		if (!file.exists()) {
			JOptionPane.showMessageDialog(null, "找不到配置文件！", "提示", JOptionPane.WARNING_MESSAGE);
		} else {
			try {
				BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(file), "utf-8"));
				String readLine = "";
				while ((readLine = br.readLine())!= null) {
					if (ADName.equals(readLine.substring(0, readLine.indexOf("\t")))) {
						// 以指定链接打开IE浏览器
						new OpenIE().open(url, cityName,batPath);
						// 截屏
						int screenWidth = ((int) java.awt.Toolkit.getDefaultToolkit().getScreenSize().width);
						int screenHeight = ((int) java.awt.Toolkit.getDefaultToolkit().getScreenSize().height);
						// 点击滚动条坐标
						float fx = Float
								.parseFloat(readLine.substring(readLine.indexOf("(") + 1, readLine.indexOf(",")));
						float fy = Float
								.parseFloat(readLine.substring(readLine.indexOf(",") + 1, readLine.indexOf(")")));
						// 释放滚动条坐标
						// float lx =
						// Float.parseFloat(readLine.substring(readLine.lastIndexOf("(")+1,readLine.lastIndexOf(",")));
						float ly = Float.parseFloat(
								readLine.substring(readLine.lastIndexOf(",") + 1, readLine.lastIndexOf(")")));
						// 点击广告坐标
						float cx = Float
								.parseFloat(readLine.substring(readLine.lastIndexOf("[") + 1, readLine.lastIndexOf("-")));
						float cy = Float
								.parseFloat(readLine.substring(readLine.lastIndexOf("-") + 1, readLine.lastIndexOf("]")));
						int x1 = (int) (fx * screenWidth);
						int y1 = (int) (fy * screenHeight);
						// int x2=(int) (lx*screenWidth);
						int y2 = (int) (ly * screenHeight);
						int x3 = (int) (cx * screenWidth);
						int y3 = (int) (cy * screenHeight);
						Robot robot = new Robot();
						robot.mouseMove(x1, y1);
						// 等待网页滚动条加载完毕
						robot.delay(2000);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseMove(x1, y2);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);

						robot.delay(4000);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseMove(x1, y1);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.delay(4000);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseMove(x1, y2);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						// 等待某些广告自动关闭
						robot.delay(5000);
						count += 1;
						// 创建屏幕截图
						CreateScreenCapture.createScreenCapture(pngPath, pngName + count);
						// 点击广告进入落地页
						robot.mouseMove(x3, y3);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.delay(60000);
						robot.delay(60000);
						robot.delay(60000);
						count += 1;
						// 创建屏幕截图
						CreateScreenCapture.createScreenCapture(pngPath, pngName + count);
						OpenIE.closeBrowse();
						robot.delay(2000);
					}
				}

			} catch (IOException | AWTException e1) {
				e1.printStackTrace();
			}
		}
	}
}
