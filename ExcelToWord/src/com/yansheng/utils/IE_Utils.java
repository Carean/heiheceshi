package com.yansheng.utils;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.IOException;

public class IE_Utils {
	/**
	 * 清除IE浏览器缓存
	 */
	public static void cleanIE(String path) {
		String command = "cmd /c start "+path;
        try {
            Runtime.getRuntime().exec(command);
        } catch (IOException e) {
            e.printStackTrace();
        }
	}
	
	/**
	 * 刷新网页
	 */
	public static void FlushIE(){
		try {
			Robot rb=new Robot();
			rb.keyPress(KeyEvent.VK_F5);
			rb.keyRelease(KeyEvent.VK_F5);
			rb.delay(1000);
			rb.keyPress(KeyEvent.VK_F5);
			rb.keyRelease(KeyEvent.VK_F5);
			rb.delay(1000);
			rb.keyPress(KeyEvent.VK_F5);
			rb.keyRelease(KeyEvent.VK_F5);
			rb.delay(1000);
			rb.keyPress(KeyEvent.VK_F5);
			rb.keyRelease(KeyEvent.VK_F5);
			rb.delay(1000);
			rb.keyPress(KeyEvent.VK_F5);
			rb.keyRelease(KeyEvent.VK_F5);
			rb.delay(1000);
		} catch (AWTException e) {
			e.printStackTrace();
		}
	}
}
