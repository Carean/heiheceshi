package com.yansheng.utils;

import java.io.File;
import java.io.FilenameFilter;
 
public class filterBypng implements FilenameFilter {
 
    @Override
    public boolean accept(File dir, String name) {
        return name.endsWith(".png");
    }
 
}