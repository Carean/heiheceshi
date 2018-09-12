package com.yansheng.utils;

import java.awt.AWTException;
import java.awt.Dimension;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;

import javax.imageio.ImageIO;

public class CreateScreenCapture {
	/**
	 * 创建全屏幕截图
	 * @param PathName  图片存储路径文件夹路径，文件名自动以系统时间编号
	 */
	public static void createScreenCapture(String PathName,String pngName) {
		try {
			//创建Robot对象
			Robot robot = new Robot();
			// 获取屏幕的尺寸
			Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
			// 获得系统属性集
			Properties props = System.getProperties();
			// 操作系统名称
			String osName = props.getProperty("os.name");
			//判断操作系统，移动鼠标至日期时间栏
			if ("Windows 10".equals(osName)) {
				robot.mouseMove(screenSize.width - 75, screenSize.height-10);
				robot.mouseMove(screenSize.width - 75, screenSize.height-15);
			} else if ("Windows 7".equals(osName)) {
				robot.mouseMove(screenSize.width - 50, screenSize.height-10);
				robot.mouseMove(screenSize.width - 50, screenSize.height-15);
			}
			//等待一秒
			robot.delay(4000);
			//调用robot类的createScreenCapture（）方法创建屏幕截图
			BufferedImage ScreenCapture = robot.createScreenCapture(new Rectangle(screenSize.width, screenSize.height));
			//获取当前系统时间
			String date = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss").format(new Date());
			//创建图片文件对象
			File file=new File(PathName+"\\"+pngName+".png");
			//如果文件不存在则创建
			if(!file.exists()){
				file.createNewFile();
			}
			//将图片写入文件（C:\\Users\\Administrator\\Desktop\\img\\1.png）
			ImageIO.write(ScreenCapture, "png",file);
			System.out.println("创建成功");
		} catch (AWTException | IOException e) {
			e.printStackTrace();
		}
	}
	
//	public static void main(String[] args) {
//		new CreateScreenCapture().createScreenCapture("C:\\Users\\Administrator\\Desktop\\img");
//	}
}
