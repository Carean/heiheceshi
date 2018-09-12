package com.yansheng.core;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.yansheng.utils.WDWUtil;

/**
 * 
 * @描述：测试excel读取 导入的jar包 poi-3.8-beta3-20110606.jar
 *               poi-ooxml-3.8-beta3-20110606.jar
 *               poi-examples-3.8-beta3-20110606.jar
 *               poi-excelant-3.8-beta3-20110606.jar
 *               poi-ooxml-schemas-3.8-beta3-20110606.jar
 *               poi-scratchpad-3.8-beta3-20110606.jar xmlbeans-2.3.0.jar
 *               dom4j-1.6.1.jar jar包官网下载地址：http://poi.apache.org/download.html
 *               下载poi-bin-3.8-beta3-20110606.zipp
 * 
 */
public class ReadExcel {

	/** 总行数 */
	private int totalRows = 0;

	/** 总列数 */
	private int totalCells = 0;

	/** 工作簿 */
	private Sheet sheet;

	/** 错误信息 */
	private String errorInfo;

	/** 构造方法 */
	public ReadExcel() {

	}

	/**
	 * @描述：得到总行数
	 * @参数：@return
	 * @返回值：int
	 */
	public int getTotalRows() {
		return totalRows;
	}

	/**
	 * @描述：得到总列数
	 * @参数：@return
	 * @返回值：int
	 */
	public int getTotalCells() {
		return totalCells;
	}

	/**
	 * @描述：得到错误信息
	 * @参数：@return
	 * @返回值：String
	 */
	public String getErrorInfo() {
		return errorInfo;
	}

	/**
	 * @描述：验证excel文件
	 * @参数：@param filePath 文件完整路径
	 * @参数：@return
	 * @返回值：boolean
	 */
	public boolean validateExcel(String filePath) {
		/** 检查文件名是否为空或者是否是Excel格式的文件 */

		if (filePath == null || !(WDWUtil.isExcel2003(filePath) || WDWUtil.isExcel2007(filePath))) {
			errorInfo = "文件名不是excel格式";
			return false;
		}

		/** 检查文件是否存在 */
		File file = new File(filePath);
		if (file == null || !file.exists()) {
			errorInfo = "文件不存在";
			return false;
		}
		return true;
	}

	/**
	 * @描述：根据文件名读取excel文件
	 * @参数：@param filePath 文件完整路径
	 * @参数：@return
	 * @返回值：List
	 */
	public List<List<String>> read(String filePath) {
		List<List<String>> dataLst = new ArrayList<List<String>>();
		InputStream is = null;
		try {
			/** 验证文件是否合法 */
			if (!validateExcel(filePath)) {
				System.out.println(errorInfo);
				return null;
			}

			/** 判断文件的类型，是2003还是2007 */
			boolean isExcel2003 = true;
			if (WDWUtil.isExcel2007(filePath)) {
				isExcel2003 = false;
			}

			/** 调用本类提供的根据流读取的方法 */
			File file = new File(filePath);
			is = new FileInputStream(file);
			dataLst = read(is, isExcel2003);
			is.close();
		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
					is = null;
					e.printStackTrace();
				}
			}
		}

		/** 返回最后读取的结果 */
		return dataLst;
	}

	/**
	 * @描述：根据流读取Excel文件
	 * @参数：@param inputStream
	 * @参数：@param isExcel2003
	 * @参数：@return
	 * @返回值：List
	 */
	public List<List<String>> read(InputStream inputStream, boolean isExcel2003) {
		List<List<String>> dataLst = null;
		try {
			/** 根据版本选择创建Workbook的方式 */
			Workbook wb = null;
			if (isExcel2003) {
				wb = new HSSFWorkbook(inputStream);
			} else {
				wb = new XSSFWorkbook(inputStream);
			}
			dataLst = read(wb);
		} catch (IOException e) {
			e.printStackTrace();
		}
		return dataLst;
	}

	/**
	 * @描述：读取数据
	 * @参数：@param Workbook
	 * @参数：@return
	 * @返回值：List<List<String>>
	 */
	private List<List<String>> read(Workbook wb) {
		List<List<String>> dataLst = new ArrayList<List<String>>();

		/** 得到第一个shell */
		this.sheet = wb.getSheetAt(0);

		/** 得到Excel的行数 */
		this.totalRows = sheet.getLastRowNum() + 1;

		/** 得到Excel的列数 */
		if (this.totalRows >= 1) {
			for (int i = 0; i < this.totalRows; i++) {
				String str = String.valueOf(sheet.getRow(i).getCell(0));
				if (str == null || str == "") {
					//System.out.println("空指针异常");
					//跳出循环
					break;
				} else if ("市场".equals(String.valueOf(sheet.getRow(i).getCell(0)))) {
					this.totalCells = sheet.getRow(i).getPhysicalNumberOfCells();
				}
			}
		}

		/** 循环Excel的行 */
		for (int r = 0; r < this.totalRows; r++) {
			Row row = sheet.getRow(r);
			if (row == null) {
				continue;
			}
			List<String> rowLst = new ArrayList<String>();

			/** 循环Excel的列 */
			for (int c = 0; c < this.getTotalCells(); c++) {
				Cell cell = row.getCell(c);
				String cellValue = "";
				if (null != cell) {
					// 以下是判断数据的类型
					switch (cell.getCellType()) {
					case HSSFCell.CELL_TYPE_NUMERIC: // 数字
						cellValue = cell.getNumericCellValue() + "";
						break;

					case HSSFCell.CELL_TYPE_STRING: // 字符串
						cellValue = cell.getStringCellValue();
						break;

					case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean
						cellValue = cell.getBooleanCellValue() + "";
						break;

					case HSSFCell.CELL_TYPE_FORMULA: // 公式
						cellValue = cell.getCellFormula() + "";
						break;

					case HSSFCell.CELL_TYPE_BLANK: // 空值
						cellValue = "";
						break;

					case HSSFCell.CELL_TYPE_ERROR: // 故障
						cellValue = "非法字符";
						break;

					default:
						cellValue = "未知类型";
						break;
					}
				}
				rowLst.add(cellValue);
			}

			/** 保存第r行的第c列 */
			dataLst.add(rowLst);
		}
		return dataLst;
	}

	/**
	 * 获取“市场”关键字所在的行数，即月份所在的行
	 * 
	 * @return
	 */
	public int getMonthRow() {
		int index = -1;
		if (this.totalRows >= 1) {
			for (int i = 0; i < this.totalRows; i++) {
				String str = String.valueOf(sheet.getRow(i).getCell(0));
				if (str == null || str == "") {
					//System.out.println("空指针异常");
					break;
				} else if ("市场".equals(String.valueOf(sheet.getRow(i).getCell(0)))) {
					index = i;
				}
			}
		}
		return index;
	}

	/**
	 * @描述：获取某月的所在列数
	 */
	public int getMonthColumn(List<List<String>> list, String mouth) {
		int index = -1;
		if (list != null) {
			List<String> cellList = list.get(getMonthRow());
			for (int j = 0; j < cellList.size(); j++) {
				if (mouth.equals(cellList.get(j)) || (mouth + "月").equals(cellList.get(j))) {
					index = j;
					//System.out.println("month所在列:" + (j));
				}
			}
		}
		return index;
	}
	
	/**
	 * @描述：获取某月某日的所在列
	 */
	public int getDayColumn(List<List<String>> list, String mouth, String day) {
		day += ".0";
		int index = -1;
		if (list != null) {
			// 获取“市场”所在行的下面一行，即每个月的日期所在行
			List<String> cellList = list.get(getMonthRow() + 1);
			int mouthColumn = getMonthColumn(list, mouth);
			//System.out.println(mouthColumn);
			//判断该月的日期应该搜索到哪一列截止，避免搜索到下一个月
			int flag=getMonthColumn(list, String.valueOf(Integer.parseInt(mouth)+1));
			if(flag<0){
				flag=cellList.size();
			}
			if (mouthColumn > 0) {
				for (int i = mouthColumn; i < flag ; i++) {
					if (day.equals(cellList.get(i))) {
						index = i;
						//System.out.println("day所在列:" + (i));
						break;
					}
				}
			} else {
				System.out.println("没有该月的排期");
			}
		}
		return index;
	}

	/**
	 * 获取某月某日值为1的数据
	 * 
	 * @param list
	 *            excel所有数据的列表
	 * @param month
	 *            月份
	 * @param day
	 *            天数
	 * @return
	 */
	public ArrayList<Map<String, String>> getPointPositionValue(List<List<String>> list, String month, String day) {
		ArrayList<Map<String, String>> maplist = new ArrayList<Map<String, String>>();
		Map<String, String> dataMap = null;
		if (list != null) {
			// 获取该月该天所在列数
			int column = getDayColumn(list, month, day);
			//System.out.println("lie:"+column);
			if (column < 0) {
				System.out.println("该天没有排期");
			} else {
				// 点位计数器
				int count = 0;
				// 统计点位个数，index = getMonthRow()+2 除去月份下面的日期行的值为1时带来的影响
				for (int index = getMonthRow() + 2; index < list.size(); index++) {
					// 获取当前行的所有数据
					List<String> cellList = list.get(index);
					if ("1.0".equals(cellList.get(column))) {
						count++;
					}
				}
				// i = getMonthRow()+2 此处为点位记录开始的行数
				for (int i = getMonthRow() + 2; i < list.size(); i++) {
					// System.out.print("第" + (i) + "行\t");
					List<String> cellList = list.get(i);
					// System.out.println(cellList);
					if ("1.0".equals(cellList.get(column))) {
//						 System.out.print("当前日期:"+month+"月"+day+"日"+"\t");
//						 System.out.print("媒体名称:"+cellList.get(1)+"\t");
//						 System.out.print("终端:"+cellList.get(2)+"\t");
//						 System.out.print("投放位置:" + cellList.get(3) + "\t");
//						 System.out.print("广告形式:" + cellList.get(4) + "\t");
//						 System.out.print("值:" + cellList.get(column - 1) +
//						 "\n");
//						 System.out.println();
						dataMap = new HashMap<String, String>();
						dataMap.put("date", month + "月" + day + "日");
						dataMap.put("mediaName", cellList.get(1));
						dataMap.put("terminal", cellList.get(2) + "端");
						dataMap.put("position", cellList.get(3));
						dataMap.put("form", cellList.get(4));
						dataMap.put("url", cellList.get(6));
						dataMap.put("count", count + "");
						maplist.add(dataMap);
					}
				}
			}
		}
		return maplist;
	}

//	/**
//	 * @描述：main测试方法
//	 * @参数：@param args
//	 * @参数：@throws Exception
//	 * @返回值：void
//	 */
//	public static void main(String[] args) throws Exception {
//		ReadExcel poi = new ReadExcel();
//		// List<List<String>> list = poi.read("d:/aaa.xls");
//		List<List<String>> list = poi.read(
//				"C:/Users/Administrator/Desktop/test/excelToword/23、2018年6-7月比亚迪华中区驻马店易车投放网络排期-双K资源具体排期（8.27-9.2）.xlsx");
//		if (list != null) {
//			for (int i = 0; i < list.size(); i++) {
//				System.out.print("第" + (i) + "行");
//				List<String> cellList = list.get(i);
//				for (int j = 0; j < cellList.size(); j++) {
//					System.out.print(" 第" + (j) + "列值：");
//					System.out.print("    " + cellList.get(j));
//				}
//				System.out.println();
//			}
//		}
//		System.out.println("总行数：" + poi.totalRows);
//		System.out.println("总列数：" + poi.totalCells);
//		System.out.println("市场:" + poi.getMonthRow());
//		System.out.println();
//		poi.getPointPositionValue(list, "8", "27");
//	}
}