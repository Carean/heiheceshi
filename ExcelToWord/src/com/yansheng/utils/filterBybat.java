package com.yansheng.utils;

import java.io.File;

import javax.swing.filechooser.FileFilter;

public class filterBybat extends FileFilter {

	@Override
	public boolean accept(File file) {
		return file.getName().endsWith(".bat");
	}

	@Override
	public String getDescription() {
		return "*.bat";  
	}

	
 
   
 
}