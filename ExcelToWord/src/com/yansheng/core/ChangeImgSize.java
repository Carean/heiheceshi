package com.yansheng.core;

import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.Shape;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;

import com.yansheng.utils.filterBypng;
import com.yansheng.utils.filterBypptx;

public class ChangeImgSize {
//	public ChangeImgSize(String pptPath){
//		// 将要读取的ppt所在目录封装成 File对象。
//		File pptdir = new File(pptPath);
//				// 通过过滤器获取目录下的所有的 .pptx 文件
//				String[] pptNameArray = pptdir.list(new filterBypptx());
//
//
//				boolean operation=false;
//				for (String pptName : pptNameArray) {
//					System.out.println(pptName);
//					try {
//						XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(new File(pptPath + "/" + pptName)));
//						List<XSLFSlide> slides = ppt.getSlides();
//						for (int index = 1; index < slides.size(); index++) {
//							XSLFSlide slide = slides.get(index);
//							for (String imgName : imgNameArray) {
//								// ppt编号
//								String pptNum = pptName.substring(0, pptName.indexOf("、"));
//								// 判断图片名是否为以当前排期编号开头且索引为当前幻灯片的索引
//								if (imgName.substring(0, imgName.indexOf(".")).equals(pptNum + "-" + index)) {
//									// 封装图片为File对象
//									File image = new File(imgPath + "/" + imgName);
//									// 读取图片
//									byte[] picture = IOUtils.toByteArray(new FileInputStream(image));
//									// 将图片添加到ppt中
//									XSLFPictureData idx = ppt.addPicture(picture, PictureData.PictureType.PNG);
//									// 在PPT中创建图片
//									XSLFPictureShape pic = slide.createPicture(idx);
//									// 获取当前幻灯片中所有组件
//									List<XSLFShape> shapes = slide.getShapes();
//									if (shapes.size()>0) {
//										// 文本内容
//										String content ="";
//										for (int i = 0; i < shapes.size(); i++) {
//											Shape shape = (Shape) shapes.get(i);
//											if (shape instanceof XSLFTextBox) {// 文本框
//												content= ((XSLFTextBox) shape).getText();
//											}
//										}
//										if (content.indexOf("APP端") >= 0||content.indexOf("app端") >= 0) {
//											// 设置图片大小及位置
//											pic.setAnchor(new Rectangle2D.Double(260, 100, 207.5, 368.5));
//											System.out.println("插入APP图片成功");
//											operation=true;
//										} else if (content.indexOf("PC端") >= 0||content.indexOf("pc端") >= 0) {
//											// 设置图片大小及位置
//											pic.setAnchor(new Rectangle2D.Double(33, 105, 654.5, 368.5));
//											System.out.println("插入PC图片成功");
//											operation=false;
//										} else if(content.indexOf("落地页") >= 0){
//											if(!operation){
//												// 设置图片大小及位置
//												pic.setAnchor(new Rectangle2D.Double(33, 105, 654.5, 368.5));
//												System.out.println("插入PC图片成功");
//											}else{
//												// 设置图片大小及位置
//												pic.setAnchor(new Rectangle2D.Double(260, 100, 207.5, 368.5));
//												System.out.println("插入APP图片成功");
//											}
//										}
//									}
//								}
//							}
//							
//							
//						}
//						// 重写ppt文件
//						FileOutputStream out = new FileOutputStream(new File(pptPath + "/" + pptName));
//						ppt.write(out);
//						out.close();
//						System.out.println("重写ppt成功");
//					} catch (IOException e) {
//						e.printStackTrace();
//					}
//				}
//	}
}
