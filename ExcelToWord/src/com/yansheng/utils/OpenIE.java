package com.yansheng.utils;

import java.io.IOException;

import javax.swing.JOptionPane;

public class OpenIE {
	/**
	 * 以制定的URL打开一个IE浏览器窗口
	 * 
	 * @param URL
	 *            要打开的链接
	 */
	public static void openIEBrowser(String URL) {
		/**
		 * 第一种方式打开IE浏览器
		 */
		// try {
		// ProcessBuilder proc = new ProcessBuilder("C:\\Program Files\\Internet
		// Explorer\\iexplore.exe", URL);
		// proc.start();
		// System.out.println("定位。"+URL);
		// } catch (Exception e) {
		// e.printStackTrace();
		// }
		/**
		 * 第二种方式打开IE浏览器
		 */
		String str = "cmd /c start iexplore " + URL;
		try {
			Runtime.getRuntime().exec(str);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
	}

	/**
	 * 关闭IE浏览器（所有窗口）
	 */
	public static void closeBrowse() {
		try {
			Runtime.getRuntime().exec("taskkill /f /im iexplore.exe");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * 以指定链接打开IE浏览器并定位
	 * 
	 * @param url
	 *            广告链接地址
	 * @param cityName
	 *            定位城市
	 */
	public void open(String url, String cityName,String batPath) {
		// 清除IE缓存
		IE_Utils.cleanIE(batPath);
		// 定位
		openIEBrowser("http://" + cityName + ".bitauto.com/cheshi/");
		try {
			// 等待网页加载
			Thread.sleep(3000);
			// 刷新网页
			IE_Utils.FlushIE();
			// 等待网页加载
			Thread.sleep(3000);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
		// 打开指定链接
		openIEBrowser(url);
		try {
			// 等待网页加载
			Thread.sleep(5000);
			// 刷新网页
			IE_Utils.FlushIE();
			// 等待网页加载
			Thread.sleep(10000);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}
}
