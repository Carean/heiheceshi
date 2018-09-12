package com.yansheng.gui;

import java.awt.Color;
import java.awt.Container;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.UIManager;

import com.yansheng.core.CreatePPT;
import com.yansheng.core.CreateWord;
import com.yansheng.core.InsertImgToPPT;
import com.yansheng.mouseSimulation.AutoSimulation;
import com.yansheng.utils.filterBybat;
import com.yansheng.utils.pngFilter;

public class GUI extends JFrame {
	// 菜单栏成员属性声明
	private JMenuBar menubar; // 菜单棒
	private JMenu helpMenu; // 帮助菜单
	private JMenuItem helpMenuItem; // 帮助菜单项
	private JPanel wordPanel, pptPanel, autoPanel;
	public static JPanel consolePanel;
	private JLabel wordTitle, inputLab, outputLab, firstpngLab, otherpngLab, dateLabel, monthLabel, dayLabel,
			fileformatLabel, pptuntilsLabel, pptfilePathLabel, imgfilePathLabel, appLabel, pcLabel, AxLabel, AyLabel,
			AwLabel, AhLabel, PxLabel, PyLabel, PwLabel, PhLabel, autoTitle, excelLab, pngLab,dateLab,monthLab,dayLab,batLab;
	private JTextField inputPath, outputPath, firstpngPath, otherpngPath, pptfilePath, imgfilePath, Ax, Ay, Aw, Ah, Px,
			Py, Pw, Ph, excelPath, pngPath,batPath;
	private JComboBox month, day, fileformat, month2, day2;
	private JButton chooseButton, chooseButton2, chooseButton3, chooseButton4,chooseButton5, choosefirstpng, chooseotherpng,
			transform, chooseppt, chooseimg, insert, startAuto;
	private String[] dayValue = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16",
			"17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31" };
	public static JTextArea console;
	public JScrollPane jsp;
	private final float standard = (float) (13.00 / 23.11);// PowerPoint软件内的高宽比

	public GUI() {
		super("拷屏辅助软件");
		setSize(745, 730);
		// 设置窗口居中
		setLocationRelativeTo(null);
		try {
			// 设置界面外观
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		} catch (Exception e) {
			System.err.println("不能被设置外观的原因:" + e);
		}
		setLayout(null);// 设置布局为空
		Container c = getContentPane(); // 创建一个内容面板
		// c.setBackground(Color.white);
		menubar = new JMenuBar(); // 创建菜单棒
		helpMenu = new JMenu("帮助");// 创建菜单
		helpMenuItem = new JMenuItem("使用说明");
		helpMenu.add(helpMenuItem);
		menubar.add(helpMenu);
		setJMenuBar(menubar); // 显示菜单栏

		helpMenuItem.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				JOptionPane.showMessageDialog(new JFrame().getContentPane(),
						"1、生成的框架文件名以选择的目录下的excel文件的文件名为准，excel的文件名应以排期编号开头以\n中文的‘、’隔开，尽量去掉文件名中不必要的部分。\n2、生成的word格式的框架为纯文本格式，不允许存储图片及其他格式，保存时弹出格式不兼容的提\n醒时选择‘否’，以最新的word格式保存一份副本即可。\n3、生成PPT框架时需提供幻灯片背景图片（png格式），一般为两张，封面一张，其他内容共用一张。\n4、使用PPT快捷插入截图功能时（可批量操作PPT文件），存储图片时需注意命名规则，按：‘排期编\n号’-‘序号’的格式，例如：排期编号为12，该排期该天共有5个点位，则需提供10张截图，命名应为：\n12-1,12-2,12-3...12-10，图片序号需对应点位截图。图片大小（单位：cm）及位置建议默认（APP端、\nPC端图片高度均默认为13cm），图片宽高比默认锁定，建议不要修改以保持图片良好的可读性。",
						"帮助", JOptionPane.INFORMATION_MESSAGE);
			}
		});

		// 创建面板
		wordPanel = new JPanel();
		wordPanel.setLayout(null);
		pptPanel = new JPanel();
		pptPanel.setLayout(null);
		autoPanel = new JPanel();
		autoPanel.setLayout(null);
		consolePanel = new JPanel();
		consolePanel.setLayout(null);

		// 标题
		wordTitle = new JLabel("拷屏文档框架生成");
		wordTitle.setFont(new Font("微软雅黑", Font.PLAIN, 20));
		wordTitle.setBounds(100, 5, 180, 30);

		// 选择文件夹组件
		inputLab = new JLabel("选择文件");
		inputPath = new JTextField();
		inputPath.setEditable(false);
		chooseButton = new JButton("选择");
		inputLab.setBounds(15, 55, 50, 30);
		inputPath.setBounds(75, 57, 190, 25);
		chooseButton.setBounds(275, 55, 60, 30);

		chooseButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				// 文件选择
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES); // 设置只选择目录
				int result = fileChooser.showOpenDialog(c); // 打开文件对话框
				if (JFileChooser.APPROVE_OPTION == result) {
					inputPath.setText(fileChooser.getSelectedFile().getPath());
				}
			}

		});

		// 选择输出路径组件
		outputLab = new JLabel("输出路径");
		outputPath = new JTextField();
		outputPath.setEditable(false);
		chooseButton2 = new JButton("选择");
		outputLab.setBounds(15, 95, 50, 30);
		outputPath.setBounds(75, 97, 190, 25);
		chooseButton2.setBounds(275, 95, 60, 30);

		chooseButton2.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				// 文件选择
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES); // 设置只选择目录
				int result = fileChooser.showOpenDialog(c); // 打开文件对话框
				if (JFileChooser.APPROVE_OPTION == result) {
					outputPath.setText(fileChooser.getSelectedFile().getPath());
				}
			}

		});

		// 输出格式
		fileformatLabel = new JLabel("输出格式");
		String[] format = { "word", "ppt" };
		fileformat = new JComboBox(format);// 下拉框
		fileformatLabel.setBounds(15, 135, 50, 30);
		fileformat.setBounds(75, 135, 190, 25);

		fileformat.addItemListener(new ItemListener() {
			@Override
			public void itemStateChanged(ItemEvent e) {
				if ("ppt".equals(fileformat.getSelectedItem().toString())) {
					wordPanel.add(firstpngLab);
					wordPanel.add(firstpngPath);
					wordPanel.add(choosefirstpng);
					wordPanel.add(otherpngLab);
					wordPanel.add(otherpngPath);
					wordPanel.add(chooseotherpng);
					dateLabel.setBounds(15, 255, 50, 30);
					month.setBounds(75, 258, 50, 25);
					monthLabel.setBounds(135, 255, 190, 30);
					day.setBounds(170, 258, 50, 25);
					dayLabel.setBounds(230, 255, 50, 30);
					transform.setBounds(265, 255, 80, 30);
					wordPanel.repaint();
				} else {
					wordPanel.remove(firstpngLab);
					wordPanel.remove(firstpngPath);
					wordPanel.remove(choosefirstpng);
					wordPanel.remove(otherpngLab);
					wordPanel.remove(otherpngPath);
					wordPanel.remove(chooseotherpng);
					dateLabel.setBounds(15, 175, 50, 30);
					month.setBounds(75, 175, 50, 25);
					monthLabel.setBounds(135, 175, 190, 30);
					day.setBounds(195, 175, 50, 25);
					dayLabel.setBounds(255, 175, 50, 30);
					transform.setBounds(125, 230, 100, 45);
					wordPanel.repaint();
				}
			}
		});

		// 选择第一张ppt背景图片路径组件
		firstpngLab = new JLabel("首页背景");
		firstpngPath = new JTextField();
		firstpngPath.setEditable(false);
		choosefirstpng = new JButton("选择");
		firstpngLab.setBounds(15, 175, 50, 30);
		firstpngPath.setBounds(75, 177, 190, 25);
		choosefirstpng.setBounds(275, 175, 60, 30);

		choosefirstpng.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				// 文件选择
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileFilter(new pngFilter());
				fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY); // 设置只选择文件
				int result = fileChooser.showOpenDialog(c); // 打开文件对话框
				if (JFileChooser.APPROVE_OPTION == result) {
					firstpngPath.setText(fileChooser.getSelectedFile().getPath());
				}
			}

		});

		// 选择其他ppt背景图片路径组件
		otherpngLab = new JLabel("其他背景");
		otherpngPath = new JTextField();
		otherpngPath.setEditable(false);
		chooseotherpng = new JButton("选择");
		otherpngLab.setBounds(15, 215, 50, 30);
		otherpngPath.setBounds(75, 217, 190, 25);
		chooseotherpng.setBounds(275, 215, 60, 30);

		chooseotherpng.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				// 文件选择
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileFilter(new pngFilter());
				fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY); // 设置只选择文件
				int result = fileChooser.showOpenDialog(c); // 打开文件对话框
				if (JFileChooser.APPROVE_OPTION == result) {
					otherpngPath.setText(fileChooser.getSelectedFile().getPath());
				}
			}

		});

		// 日期面板
		dateLabel = new JLabel("选择日期");
		monthLabel = new JLabel("月");
		dayLabel = new JLabel("日");
		String[] province = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12" };
		month = new JComboBox(province);// 下拉框
		month.setMaximumRowCount(5);
		day = new JComboBox(dayValue);// 下拉框
		day.setMaximumRowCount(5);
		dateLabel.setBounds(15, 175, 50, 30);
		month.setBounds(75, 178, 50, 25);
		monthLabel.setBounds(135, 175, 190, 30);
		day.setBounds(170, 178, 50, 25);
		dayLabel.setBounds(230, 175, 50, 30);

		// 月份改变监听
		month.addItemListener(new ItemListener() {

			@Override
			public void itemStateChanged(ItemEvent e) {
				if (e.getStateChange() == ItemEvent.SELECTED) {
					String monthValue = month.getSelectedItem().toString();
					switch (monthValue) {
					case "1":
					case "3":
					case "5":
					case "7":
					case "8":
					case "10":
					case "12":
						day.removeAllItems();
						for (String string : dayValue) {
							day.addItem(string);
						}
						break;

					case "2":
						day.removeAllItems();
						String[] item = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14",
								"15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28",
								"29" };
						for (String string : item) {
							day.addItem(string);
						}
						break;

					case "4":
					case "6":
					case "9":
					case "11":
						day.removeAllItems();
						String[] items = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14",
								"15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28",
								"29", "30" };
						for (String string : items) {
							day.addItem(string);
						}
						break;

					default:
						break;
					}
				}
			}
		});

		// 开始转换按钮
		transform = new JButton("开始转换");
		transform.setBounds(125, 230, 100, 45);
		transform.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				if ("".equals(inputPath.getText().trim()) || "".equals(outputPath.getText().trim())) {
					JOptionPane.showMessageDialog(new JFrame().getContentPane(), "转换失败！请先选择文件路径", "错误",
							JOptionPane.ERROR_MESSAGE);
				} else {
					if ("word".equals(fileformat.getSelectedItem().toString())) {
						boolean createWord = new CreateWord().CreateWord(inputPath.getText(), outputPath.getText(),
								month.getSelectedItem().toString(), day.getSelectedItem().toString());
						if (createWord) {
							JOptionPane.showMessageDialog(new JFrame().getContentPane(), "转换完成！", "提示",
									JOptionPane.INFORMATION_MESSAGE);
						} else {
							JOptionPane.showMessageDialog(new JFrame().getContentPane(), "转换失败！请确保未打开需要转换的文件！", "错误",
									JOptionPane.ERROR_MESSAGE);
						}
					} else {
						if ("".equals(firstpngPath.getText().trim()) || "".equals(otherpngPath.getText().trim())) {
							JOptionPane.showMessageDialog(new JFrame().getContentPane(), "转换失败！请先选择文件路径", "错误",
									JOptionPane.ERROR_MESSAGE);
						} else {
							boolean createPPT = new CreatePPT().CreatePPT(inputPath.getText(), outputPath.getText(),
									firstpngPath.getText(), otherpngPath.getText(), month.getSelectedItem().toString(),
									day.getSelectedItem().toString());
							if (createPPT) {
								JOptionPane.showMessageDialog(new JFrame().getContentPane(), "转换完成！", "提示",
										JOptionPane.INFORMATION_MESSAGE);
							} else {
								JOptionPane.showMessageDialog(new JFrame().getContentPane(), "转换失败！请确保未打开需要转换的文件！",
										"错误", JOptionPane.ERROR_MESSAGE);
							}
						}
					}
				}
			}
		});

		// 设置word面板
		wordPanel.setBackground(Color.white);
		wordPanel.setBounds(0, 0, 360, 290);
		wordPanel.add(wordTitle);
		wordPanel.add(inputLab);
		wordPanel.add(inputPath);
		wordPanel.add(chooseButton);
		wordPanel.add(outputLab);
		wordPanel.add(outputPath);
		wordPanel.add(chooseButton2);
		wordPanel.add(fileformatLabel);
		wordPanel.add(fileformat);
		wordPanel.add(dateLabel);
		wordPanel.add(month);
		wordPanel.add(monthLabel);
		wordPanel.add(day);
		wordPanel.add(dayLabel);
		wordPanel.add(transform);

		// 标题
		pptuntilsLabel = new JLabel("PPT快捷插入截图");
		pptuntilsLabel.setFont(new Font("微软雅黑", Font.PLAIN, 20));
		pptuntilsLabel.setBounds(100, 5, 180, 30);

		// 选择文件夹组件
		pptfilePathLabel = new JLabel("PPT路径");
		pptfilePath = new JTextField(30);
		pptfilePath.setEditable(false);
		chooseppt = new JButton("选择");
		pptfilePathLabel.setBounds(15, 55, 50, 30);
		pptfilePath.setBounds(75, 57, 190, 25);
		chooseppt.setBounds(275, 55, 60, 30);

		chooseppt.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				// 文件选择
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES); // 设置只选择目录
				int result = fileChooser.showOpenDialog(c); // 打开文件对话框
				if (JFileChooser.APPROVE_OPTION == result) {
					pptfilePath.setText(fileChooser.getSelectedFile().getPath());
				}
			}

		});

		// 选择输出路径组件
		imgfilePathLabel = new JLabel("图片路径");
		imgfilePath = new JTextField(30);
		imgfilePath.setEditable(false);
		chooseimg = new JButton("选择");
		imgfilePathLabel.setBounds(15, 95, 50, 30);
		imgfilePath.setBounds(75, 97, 190, 25);
		chooseimg.setBounds(275, 95, 60, 30);

		chooseimg.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				// 文件选择
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES); // 设置只选择目录
				int result = fileChooser.showOpenDialog(c); // 打开文件对话框
				if (JFileChooser.APPROVE_OPTION == result) {
					imgfilePath.setText(fileChooser.getSelectedFile().getPath());
				}
			}

		});

		appLabel = new JLabel("APP端图片");
		AxLabel = new JLabel("x:");
		Ax = new JTextField("260");
		AyLabel = new JLabel("y:");
		Ay = new JTextField("100");
		AwLabel = new JLabel("w:");
		Aw = new JTextField("7.31");
		AhLabel = new JLabel("h:");
		Ah = new JTextField("13.00");
		appLabel.setBounds(8, 135, 60, 30);
		AxLabel.setBounds(70, 135, 40, 25);
		AyLabel.setBounds(140, 135, 40, 25);
		AwLabel.setBounds(210, 135, 40, 25);
		AhLabel.setBounds(280, 135, 40, 25);
		Ax.setBounds(85, 135, 40, 25);
		Ay.setBounds(155, 135, 40, 25);
		Aw.setBounds(225, 135, 40, 25);
		Ah.setBounds(295, 135, 40, 25);

		pcLabel = new JLabel("PC端图片");
		PxLabel = new JLabel("x:");
		Px = new JTextField("33");
		PyLabel = new JLabel("y:");
		Py = new JTextField("105");
		PwLabel = new JLabel("w:");
		Pw = new JTextField("23.11");
		PhLabel = new JLabel("h:");
		Ph = new JTextField("13.00");
		pcLabel.setBounds(15, 175, 50, 30);
		PxLabel.setBounds(70, 175, 40, 25);
		PyLabel.setBounds(140, 175, 40, 25);
		PwLabel.setBounds(210, 175, 40, 25);
		PhLabel.setBounds(280, 175, 40, 25);
		Px.setBounds(85, 175, 40, 25);
		Py.setBounds(155, 175, 40, 25);
		Pw.setBounds(225, 175, 40, 25);
		Ph.setBounds(295, 175, 40, 25);

		Aw.addKeyListener(new KeyListener() {
			public void keyTyped(KeyEvent e) {
			}

			public void keyReleased(KeyEvent e) {
				if ("".equals(Aw.getText().trim())) {
					Aw.setText("7.30");
					Ah.setText(String.format("%.2f", (Float.parseFloat(Aw.getText())) / standard));
				} else {
					Ah.setText(String.format("%.2f", (Float.parseFloat(Aw.getText())) / standard));
				}
			}

			public void keyPressed(KeyEvent e) {
			}
		});

		Ah.addKeyListener(new KeyListener() {
			public void keyTyped(KeyEvent e) {
			}

			public void keyReleased(KeyEvent e) {
				if ("".equals(Ah.getText().trim())) {
					Ah.setText("13.00");
					Aw.setText(String.format("%.2f", (Float.parseFloat(Ah.getText())) * standard));
				} else {
					Aw.setText(String.format("%.2f", (Float.parseFloat(Ah.getText())) * standard));
				}
			}

			public void keyPressed(KeyEvent e) {
			}
		});

		Pw.addKeyListener(new KeyListener() {
			public void keyTyped(KeyEvent e) {
			}

			public void keyReleased(KeyEvent e) {
				if ("".equals(Pw.getText().trim())) {
					Pw.setText("23.11");
					Ph.setText(String.format("%.2f", (Float.parseFloat(Pw.getText())) * standard));
				} else {
					Ph.setText(String.format("%.2f", (Float.parseFloat(Pw.getText())) * standard));
				}
			}

			public void keyPressed(KeyEvent e) {
			}
		});

		Ph.addKeyListener(new KeyListener() {
			public void keyTyped(KeyEvent e) {
			}

			public void keyReleased(KeyEvent e) {
				if ("".equals(Ph.getText().trim())) {
					Ph.setText("13.00");
					Pw.setText(String.format("%.2f", (Float.parseFloat(Ph.getText())) / standard));
				} else {
					Pw.setText(String.format("%.2f", (Float.parseFloat(Ph.getText())) / standard));
				}
			}

			public void keyPressed(KeyEvent e) {
			}
		});

		// 开始转换按钮
		insert = new JButton("插入图片");
		insert.setBounds(125, 230, 100, 45);
		insert.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				if ("".equals(pptfilePath.getText().trim()) || "".equals(imgfilePath.getText().trim())) {
					JOptionPane.showMessageDialog(new JFrame().getContentPane(), "插入失败！请先选择文件路径", "错误",
							JOptionPane.ERROR_MESSAGE);
				} else {
					boolean insertImage = new InsertImgToPPT().InsertImage(pptfilePath.getText(), imgfilePath.getText(),
							Float.parseFloat(Ax.getText()), Float.parseFloat(Ay.getText()),
							Float.parseFloat(Aw.getText()), Float.parseFloat(Ah.getText()),
							Float.parseFloat(Px.getText()), Float.parseFloat(Py.getText()),
							Float.parseFloat(Pw.getText()), Float.parseFloat(Ph.getText()));
					if (insertImage) {
						JOptionPane.showMessageDialog(new JFrame().getContentPane(), "插入完成！", "提示",
								JOptionPane.INFORMATION_MESSAGE);
					} else {
						JOptionPane.showMessageDialog(new JFrame().getContentPane(), "插入失败！请确保未打开需要插入的文件！", "错误",
								JOptionPane.ERROR_MESSAGE);
					}

				}
			}
		});

		pptPanel.setBackground(Color.white);
		pptPanel.setBounds(380, 0, 360, 290);
		pptPanel.add(pptuntilsLabel);
		pptPanel.add(pptfilePathLabel);
		pptPanel.add(pptfilePath);
		pptPanel.add(chooseppt);
		pptPanel.add(imgfilePathLabel);
		pptPanel.add(imgfilePath);
		pptPanel.add(chooseimg);
		pptPanel.add(appLabel);
		pptPanel.add(AxLabel);
		pptPanel.add(AyLabel);
		pptPanel.add(AwLabel);
		pptPanel.add(AhLabel);
		pptPanel.add(Ax);
		pptPanel.add(Ay);
		pptPanel.add(Aw);
		pptPanel.add(Ah);
		pptPanel.add(pcLabel);
		pptPanel.add(PxLabel);
		pptPanel.add(PyLabel);
		pptPanel.add(PwLabel);
		pptPanel.add(PhLabel);
		pptPanel.add(Px);
		pptPanel.add(Py);
		pptPanel.add(Pw);
		pptPanel.add(Ph);
		pptPanel.add(insert);

		// 标题
		autoTitle = new JLabel("易车自动拷屏");
		autoTitle.setFont(new Font("微软雅黑", Font.PLAIN, 20));
		autoTitle.setBounds(305, 5, 180, 30);

		// 选择文件夹组件
		excelLab = new JLabel("选择文件");
		excelPath = new JTextField();
		excelPath.setEditable(false);
		chooseButton3 = new JButton("选择");
		excelLab.setBounds(15, 55, 50, 30);
		excelPath.setBounds(75, 57, 190, 25);
		chooseButton3.setBounds(275, 55, 60, 30);

		chooseButton3.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				// 文件选择
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES); // 设置只选择目录
				int result = fileChooser.showOpenDialog(c); // 打开文件对话框
				if (JFileChooser.APPROVE_OPTION == result) {
					excelPath.setText(fileChooser.getSelectedFile().getPath());
				}
			}

		});

		// 选择输出路径组件
		pngLab = new JLabel("截图存储");
		pngPath = new JTextField();
		pngPath.setEditable(false);
		chooseButton4 = new JButton("选择");
		pngLab.setBounds(395, 55, 50, 30);
		pngPath.setBounds(455, 57, 190, 25);
		chooseButton4.setBounds(655, 55, 60, 30);

		chooseButton4.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				// 文件选择
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES); // 设置只选择目录
				int result = fileChooser.showOpenDialog(c); // 打开文件对话框
				if (JFileChooser.APPROVE_OPTION == result) {
					pngPath.setText(fileChooser.getSelectedFile().getPath());
				}
			}

		});

		// 日期面板
		dateLab = new JLabel("选择日期");
		monthLab = new JLabel("月");
		dayLab = new JLabel("日");
		month2 = new JComboBox(province);// 下拉框
		month2.setMaximumRowCount(5);
		day2 = new JComboBox(dayValue);// 下拉框
		day2.setMaximumRowCount(5);
		dateLab.setBounds(15, 105, 50, 30);
		month2.setBounds(75, 108, 50, 25);
		monthLab.setBounds(135, 105, 190, 30);
		day2.setBounds(170, 108, 50, 25);
		dayLab.setBounds(230, 105, 50, 30);

		// 月份改变监听
		month2.addItemListener(new ItemListener() {

			@Override
			public void itemStateChanged(ItemEvent e) {
				if (e.getStateChange() == ItemEvent.SELECTED) {
					String monthValue = month2.getSelectedItem().toString();
					switch (monthValue) {
					case "1":
					case "3":
					case "5":
					case "7":
					case "8":
					case "10":
					case "12":
						day2.removeAllItems();
						for (String string : dayValue) {
							day2.addItem(string);
						}
						break;

					case "2":
						day2.removeAllItems();
						String[] item = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14",
								"15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28",
								"29" };
						for (String string : item) {
							day2.addItem(string);
						}
						break;

					case "4":
					case "6":
					case "9":
					case "11":
						day2.removeAllItems();
						String[] items = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14",
								"15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28",
								"29", "30" };
						for (String string : items) {
							day2.addItem(string);
						}
						break;

					default:
						break;
					}
				}
			}
		});
		
		// 选择bat文件路径组件
		batLab = new JLabel("清理缓存");
		batPath = new JTextField();
		batPath.setEditable(false);
		chooseButton5 = new JButton("选择");
		batLab.setBounds(395, 105, 50, 30);
		batPath.setBounds(455, 107, 190, 25);
		chooseButton5.setBounds(655, 105, 60, 30);

		chooseButton5.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				// 文件选择
				JFileChooser fileChooser = new JFileChooser();
				//fileChooser.addChoosableFileFilter(new filterBybat());
				fileChooser.setFileFilter(new filterBybat());
				fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY); // 设置只选择文件
				int result = fileChooser.showOpenDialog(c); // 打开文件对话框
				if (JFileChooser.APPROVE_OPTION == result) {
					batPath.setText(fileChooser.getSelectedFile().getPath());
				}
			}

		});
		
		
		startAuto=new JButton("开始自动拷屏");
		startAuto.setBounds(310, 140, 120, 30);
		startAuto.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				if ("".equals(excelPath.getText().trim()) || "".equals(pngPath.getText().trim())|| "".equals(batPath.getText().trim())) {
					JOptionPane.showMessageDialog(new JFrame().getContentPane(), "请先选择文件路径", "错误",
							JOptionPane.ERROR_MESSAGE);
				} else {
					//最小化窗口
					setExtendedState(JFrame.ICONIFIED);
					new AutoSimulation().startSimulation(excelPath.getText(), pngPath.getText(),batPath.getText(), month2.getSelectedItem().toString(), day2.getSelectedItem().toString());
				}
			}
		});

		autoPanel.setBackground(Color.white);
		autoPanel.setBounds(0, 310, 745, 180);
		autoPanel.add(autoTitle);
		autoPanel.add(excelLab);
		autoPanel.add(excelPath);
		autoPanel.add(chooseButton3);
		autoPanel.add(pngLab);
		autoPanel.add(pngPath);
		autoPanel.add(chooseButton4);
		autoPanel.add(dateLab);
		autoPanel.add(month2);
		autoPanel.add(monthLab);
		autoPanel.add(day2);
		autoPanel.add(dayLab);
		autoPanel.add(batLab);
		autoPanel.add(batPath);
		autoPanel.add(chooseButton5);
		autoPanel.add(startAuto);

		console = new JTextArea();
		// console.setLineWrap(true);//设置自动换行
		// console.setBounds(0, 10, 710, 150);
		jsp = new JScrollPane(console);
		// jsp.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
		jsp.setBounds(0, 0, 705, 150);
		consolePanel.setBounds(15, 510, 705, 150);
		// consolePanel.setBackground(Color.white);
		consolePanel.add(jsp);

		c.add(wordPanel);
		c.add(pptPanel);
		c.add(autoPanel);
		c.add(consolePanel);
		setVisible(true);
		setResizable(false);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	}

	public static void main(String[] args) {
		new GUI();
	}
}
