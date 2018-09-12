package com.yansheng.core;

import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.Shape;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;

import com.yansheng.gui.GUI;
import com.yansheng.utils.filterBypng;
import com.yansheng.utils.filterBypptx;

public class InsertImgToPPT {
	private final float flag=(float) (368.5/13);//一厘米相当于多少像素，368.5在PowerPoint内为13cm

	/**
	 * 
	 * @param pptPath	ppt路径
	 * @param imgPath	图片路径
	 * @param Ax    app端图片X轴
	 * @param Ay	app端图片Y轴
	 * @param Aw	app端图片宽（单位：cm）
	 * @param Ah	app端图片高（单位：cm）
	 * @param Px	pc端图片X轴
	 * @param Py	pc端图片Y轴
	 * @param Pw	pc端图片宽（单位：cm）
	 * @param Ph	pc端图片高（单位：cm）
	 * @return
	 */
	public boolean InsertImage(String pptPath, String imgPath,float Ax,float Ay,float Aw,float Ah,float Px,float Py,float Pw,float Ph) {
		// 将要读取的ppt所在目录封装成 File对象。
		File pptdir = new File(pptPath);
		// 通过过滤器获取目录下的所有的 .pptx 文件
		String[] pptNameArray = pptdir.list(new filterBypptx());
		if(pptNameArray.length<=0){
			JOptionPane.showMessageDialog(new JFrame().getContentPane(), "当前目录下没有.pptx文件", "错误",
					JOptionPane.ERROR_MESSAGE);
			return false;
		}
		// 将要读取的png图片所在目录封装成 File对象。
		File imgdir = new File(imgPath);
		// 通过过滤器获取目录下的所有的 .png 文件
		String[] imgNameArray = imgdir.list(new filterBypng());
		if(pptNameArray.length<=0){
			JOptionPane.showMessageDialog(new JFrame().getContentPane(), "当前目录下没有.png文件", "错误",
					JOptionPane.ERROR_MESSAGE);
			return false;
		}
		boolean operation=false;
		for (String pptName : pptNameArray) {
			GUI.console.append(pptName+"\n");
			GUI.console.paintImmediately(GUI.console.getBounds());//立即刷新文本域的值
			GUI.consolePanel.paintImmediately(GUI.consolePanel.getBounds());
			System.out.println(pptName);
			try {
				XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(new File(pptPath + "/" + pptName)));
				List<XSLFSlide> slides = ppt.getSlides();
				for (int index = 1; index < slides.size(); index++) {
					XSLFSlide slide = slides.get(index);
					for (String imgName : imgNameArray) {
						// ppt编号
						String pptNum = pptName.substring(0, pptName.indexOf("、"));
						// 判断图片名是否为以当前排期编号开头且索引为当前幻灯片的索引
						if (imgName.substring(0, imgName.indexOf(".")).equals(pptNum + "-" + index)) {
							// 封装图片为File对象
							File image = new File(imgPath + "/" + imgName);
							// 读取图片
							byte[] picture = IOUtils.toByteArray(new FileInputStream(image));
							// 将图片添加到ppt中
							XSLFPictureData idx = ppt.addPicture(picture, PictureData.PictureType.PNG);
							// 在PPT中创建图片
							XSLFPictureShape pic = slide.createPicture(idx);
							// 获取当前幻灯片中所有组件
							List<XSLFShape> shapes = slide.getShapes();
							if (shapes.size()>0) {
								// 文本内容
								String content ="";
								for (int i = 0; i < shapes.size(); i++) {
									Shape shape = (Shape) shapes.get(i);
									if (shape instanceof XSLFTextBox) {// 文本框
										content= ((XSLFTextBox) shape).getText();
									}
								}
								if (content.indexOf("APP端") >= 0||content.indexOf("app端") >= 0) {
									// 设置图片大小及位置
									//pic.setAnchor(new Rectangle2D.Double(260, 100, 207.5, 368.5));
									pic.setAnchor(new Rectangle2D.Double(Ax, Ay, Aw*flag, Ah*flag));
									GUI.console.append("插入APP图片成功\n");
									GUI.console.paintImmediately(GUI.console.getBounds());//立即刷新文本域的值
									GUI.consolePanel.paintImmediately(GUI.consolePanel.getBounds());
									System.out.println("插入APP图片成功");
									operation=true;
								} else if (content.indexOf("PC端") >= 0||content.indexOf("pc端") >= 0) {
									// 设置图片大小及位置
									//pic.setAnchor(new Rectangle2D.Double(33, 105, 654.5, 368.5));
									pic.setAnchor(new Rectangle2D.Double(Px, Py, Pw*flag, Ph*flag));
									GUI.console.append("插入PC图片成功\n");
									GUI.console.paintImmediately(GUI.console.getBounds());//立即刷新文本域的值
									GUI.consolePanel.paintImmediately(GUI.consolePanel.getBounds());
									System.out.println("插入PC图片成功");
									operation=false;
								} else if(content.indexOf("落地页") >= 0){
									if(!operation){
										// 设置图片大小及位置
										//pic.setAnchor(new Rectangle2D.Double(33, 105, 654.5, 368.5));
										pic.setAnchor(new Rectangle2D.Double(Px, Py, Pw*flag, Ph*flag));
										GUI.console.append("插入PC图片成功\n");
										GUI.console.paintImmediately(GUI.console.getBounds());//立即刷新文本域的值
										GUI.consolePanel.paintImmediately(GUI.consolePanel.getBounds());
										System.out.println("插入PC图片成功");
									}else{
										// 设置图片大小及位置
										//pic.setAnchor(new Rectangle2D.Double(260, 100, 207.5, 368.5));
										pic.setAnchor(new Rectangle2D.Double(Ax, Ay, Aw*flag, Ah*flag));
										GUI.console.append("插入APP图片成功\n");
										GUI.console.paintImmediately(GUI.console.getBounds());//立即刷新文本域的值
										GUI.consolePanel.paintImmediately(GUI.consolePanel.getBounds());
										System.out.println("插入APP图片成功");
									}
								}
							}
						}
					}
					
					
				}
				// 重写ppt文件
				FileOutputStream out = new FileOutputStream(new File(pptPath + "/" + pptName));
				ppt.write(out);
				out.close();
				GUI.console.setText(GUI.console.getText()+"重写ppt成功\n\n");
				System.out.println("重写ppt成功");
			} catch (IOException e) {
				e.printStackTrace();
				JOptionPane.showMessageDialog(new JFrame().getContentPane(), "插入失败！请确保未打开需要操作的文件！", "错误",
						JOptionPane.ERROR_MESSAGE);
				return false;
			}
		}
		return true;
	}

//	public static void main(String[] args) {
//		// //转义ppt路径
//		// String pptPath="C:\\Users\\Administrator\\Desktop\\test\\ppt";
//		// String newpptPath=pptPath.replace("\\\\","/");
//		// //转义图片路径
//		// String path="C:\\Users\\Administrator\\Desktop\\test\\img";
//		// String newPath=path.replace("\\\\","/");
//		boolean insertImage = new InsertImgToPPT().InsertImage("C:/Users/Administrator/Desktop/test/ppt",
//				"C:/Users/Administrator/Desktop/test/img");
//		System.out.println(insertImage);
//	}
}
