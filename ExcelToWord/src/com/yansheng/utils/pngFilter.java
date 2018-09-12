package com.yansheng.utils;

import java.io.File;

import javax.swing.filechooser.FileFilter;

public class pngFilter extends FileFilter {

	@Override
	public boolean accept(File f) {
		if (f.isDirectory()) {
			return true;
		}
		String filename=f.getName();
		String extension = filename.substring(filename.indexOf(".")+1,filename.length());
		if (extension != null) {
			if (extension.equals("png")) {
				return true;
			} else {
				return false;
			}
		}
		return false;
	}

	@Override
	public String getDescription() {
		return "图片文件(*.png)";
	}

}
