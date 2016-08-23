package com.yht.exceltemplate;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import jxl.Cell;
import jxl.CellType;
import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.write.Label;
import jxl.write.WritableCellFeatures;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class ExcelTemplate {
	Map<String,Object> keyMap = new HashMap<String, Object>();

	public ExcelTemplate(Map<String,Object> map){
		keyMap = map;
	}




	private Workbook readTemplate(InputStream inputStream) throws Exception {
		Workbook rwb = Workbook.getWorkbook(inputStream);
		return rwb;
	}

	private List<Map> getList(String listName, String itemName, int listSize) throws Exception {
		itemName = itemName.trim();
		List<Map> list = new ArrayList<Map>();

//        Map m = new HashMap();
//        m.put(itemName+".名称", "张三");
//        m.put(itemName+".年龄", "12");
//        list.add(m);
//        Map m1 = new HashMap();
//        m1.put(itemName+".名称", "李四");
//        m1.put(itemName+".年龄", "13");
//        list.add(m1);
		if(listName != null && !listName.equals("")){
			List<Map<String,Object>> listold = (List)keyMap.get(listName);
			for(Map<String,Object> m : listold){
				Map newMap = new HashMap();
				for(String key : m.keySet()){
					newMap.put(itemName + "." + key, m.get(key));
				}
				list.add(newMap);
			}
		}
		return list;
	}

	public void printTemplate(InputStream is,OutputStream outputStream){
		try {
			analysisTemplate(is,outputStream,new ArrayList<String>(keyMap.keySet()), keyMap);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void analysisTemplate(InputStream is,OutputStream outputStream,List<String> itemList, Map itemValueMap) throws Exception {
		Workbook rwb = null;
		rwb = readTemplate(is);
		Sheet readSheet = rwb.getSheet(0);

		List<Map> beginList = new ArrayList<Map>();
		List<Map> endList = new ArrayList<Map>();
		List<Map> removeRowList = new ArrayList<Map>();

		int rows = readSheet.getRows();//获取工作表中的总行数
		int columns = readSheet.getColumns();//获取工作表中的总列数
		//获取begin和end的位置，由于不支持嵌套，所以第一个begin和第一个end一定是匹配的
		for (int i = 0; i < rows; i++) {
			for (int j = 0; j < columns; j++) {
				Cell oCell = readSheet.getCell(j, i);//需要注意的是这里的getCell方法的参数，第一个是指定第几列，第二个参数才是指定第几行
				if (!oCell.getType().equals(CellType.EMPTY)) {
					String contents = oCell.getContents();
					if (contents != null) {
						if (contents.indexOf("$={begin") != -1) {
							Map m = new HashMap();
							m.put("row", oCell.getRow());
							m.put("column", oCell.getColumn());
							m.put("content", contents);
							beginList.add(m);
						} else if (contents.indexOf("$={end}") != -1) {
							Map m = new HashMap();
							m.put("row", oCell.getRow());
							m.put("column", oCell.getColumn());
							endList.add(m);
						}
					}
				}
			}
		}
		List<Map> eachValues = new ArrayList<Map>();
		List<Map> formatMergedList = new ArrayList<Map>();
		List<Map> removeFormatMergedList = new ArrayList<Map>();
		List<Map> rowHeightList = new ArrayList<Map>();
		Set<Integer> formatrowPageBreakList = new HashSet<Integer>();
		Range[] ranges = readSheet.getMergedCells();
		if (beginList.size() > 0 && endList.size() > 0) {
			Map beginMap = beginList.get(0);
			int beginRow = (Integer) beginMap.get("row");
			for (Range range : ranges) {
				if (beginRow < range.getTopLeft().getRow()) {
					Map removeMeargedMap = new HashMap();
					removeMeargedMap.put("left", range.getTopLeft().getColumn());
					removeMeargedMap.put("top", range.getTopLeft().getRow());
					removeMeargedMap.put("right", range.getBottomRight().getColumn());
					removeMeargedMap.put("down", range.getBottomRight().getRow());
					removeFormatMergedList.add(removeMeargedMap);
				}
			}
		}
		int[] rowPageBreaks = readSheet.getRowPageBreaks();
		int lastEachRow = 0;
		for (int i = 0; i < beginList.size(); i++) {
			if (i < endList.size()) {
				Map beginMap = beginList.get(i);
				Map endMap = endList.get(i);
				String contents = beginMap.get("content").toString();
				int beginRow = (Integer) beginMap.get("row");
				int endRow = (Integer) endMap.get("row");
				if (endRow > beginRow) {
					int endColumn = (Integer) endMap.get("column");
					String[] props = contents.split(":");
					if (props.length == 0) {
						beginList.get(i).put("error", true);
						break;
					}
					if (props.length <= 1) {
						beginList.get(i).put("error", true);
						break;
					}
					if (props[1].split("_").length != 2) {
						beginList.get(i).put("error", true);
						break;
					}
					String listName = props[1].split("_")[0];
					String itemName = props[1].split("_")[1];
					int listSize = Integer.MAX_VALUE;
					if (props.length >= 3) {
						try {
							String sizeString = props[2];
							sizeString = sizeString.replace("}", "");
							listSize = Integer.parseInt(sizeString);
						} catch (Exception e) {
							//模版格式错误，这里应该是数字
						}
					} else {
						itemName = itemName.replace("}", "");
					}
					List<Map> rangeInEach = new ArrayList<Map>();
					for (Range range : ranges) {
						if (range.getTopLeft().getRow() > beginRow && range.getBottomRight().getRow() < endRow) {
							Map rangeInEachMap = new HashMap();
							rangeInEachMap.put("betweenRowCount", endRow - beginRow - 1);
							rangeInEachMap.put("range", range);
							rangeInEach.add(rangeInEachMap);
						}
					}
					List<Integer> rowPageBreakInEach = new ArrayList<Integer>();
					if (rowPageBreaks != null) {
						for (int rowPageBreak : rowPageBreaks) {
							if (rowPageBreak > beginRow && rowPageBreak <= endRow) {
								rowPageBreakInEach.add(rowPageBreak);
							}
						}
					}
					List<Map> dataList = getList(listName, itemName, listSize);//其他list从数据库查询  查询赋值
					if (dataList.size() == 0 && i == 0) {
						for (int r = beginRow; r < endRow; r++) {
							Map centerRowMap = new HashMap();
							centerRowMap.put("row", beginRow);
						}
					}

					int betweenRowCount = endRow - beginRow - 1;
					List<Map> tableFirstRows = new ArrayList<Map>();
					if (i == 0) {
						for (int row = beginRow + 1; row <= endRow; row++) {
							for (int col = 0; col < columns; col++) {
								Cell oCell = readSheet.getCell(col, row);
								Map oCellMap = new HashMap();
								oCellMap.put("cell", oCell);
								oCellMap.put("row", oCell.getRow());
								tableFirstRows.add(oCellMap);
							}
						}
					} else {
						Map lastEndMap = endList.get(i - 1);
						int lastEndRow = (Integer) lastEndMap.get("row");
						if (beginList.size() > 1 && endList.size() > 1 && i > 0) {
							if (lastEachRow == 0) {
								lastEachRow = lastEndRow;
							}
							if (lastEachRow > 0) {
								int middleRow = lastEachRow - 1;
								for (Range range : ranges) {
									if (range.getTopLeft().getRow() >= lastEndRow + 1 && range.getBottomRight().getRow() < beginRow) {
										Map rangeMap = new HashMap();
										rangeMap.put("left", range.getTopLeft().getColumn());
										rangeMap.put("top", middleRow + range.getTopLeft().getRow() - lastEndRow - 1);
										rangeMap.put("right", range.getBottomRight().getColumn());
										rangeMap.put("down", middleRow + range.getBottomRight().getRow() - lastEndRow - 1);
										formatMergedList.add(rangeMap);
									}
								}
								Map<Integer, Map> rowHeightMapMap = new HashMap<Integer, Map>();
								for (int row = lastEndRow + 1; row < beginRow; row++) {
									for (int column = 0; column < columns; column++) {
										Cell oCell = readSheet.getCell(column, row);
										Map eachCellMap = new HashMap();
										eachCellMap.put("cell", oCell);
										eachCellMap.put("row", middleRow);
										eachCellMap.put("column", oCell.getColumn());
										eachCellMap.put("content", oCell.getContents());
										eachValues.add(eachCellMap);
										if (rowHeightMapMap.get(middleRow) == null) {
											int rowHeight = readSheet.getRowHeight(oCell.getRow());
											Map rowHeightMap = new HashMap();
											rowHeightMap.put("row", middleRow);
											rowHeightMap.put("height", rowHeight);
											rowHeightMapMap.put(middleRow, rowHeightMap);
										}
									}
									middleRow++;
								}

								Set<Integer> rowHeightSet = rowHeightMapMap.keySet();

								for (Integer rowHeight : rowHeightSet) {
									Map rowHeightMap = rowHeightMapMap.get(rowHeight);
									if (rowHeightMap != null) {
										rowHeightList.add(rowHeightMap);
									}
								}
							}
						}
						lastEachRow = lastEachRow + (beginRow - lastEndRow - 1);
						for (int row = beginRow + 1; row <= endRow; row++) {
							for (int col = 0; col < columns; col++) {
								Cell oCell = readSheet.getCell(col, row);
								Map oCellMap = new HashMap();
								oCellMap.put("cell", oCell);
								oCellMap.put("row", oCell.getRow() - (beginRow + 1) + lastEachRow);
								tableFirstRows.add(oCellMap);
							}
						}
					}
					if (dataList.size() == 0 && i == 0) {
						for (int r = beginRow; r < endRow - 1; r++) {
							Map centerRowMap = new HashMap();
							centerRowMap.put("row", beginRow);
							removeRowList.add(centerRowMap);
						}
					}
					for (int index = 0; index < dataList.size(); index++) {
						Map dataMap = dataList.get(index);
						for (Map rangeInEachMap : rangeInEach) {
							Range range = (Range) rangeInEachMap.get("range");
							Integer betweenRowCountEach = (Integer) rangeInEachMap.get("betweenRowCount");
							Map rangeMap = new HashMap();
							if (i == 0) {
								rangeMap.put("left", range.getTopLeft().getColumn());
								rangeMap.put("top", range.getTopLeft().getRow() + index * betweenRowCount - 1);
								rangeMap.put("right", range.getBottomRight().getColumn());
								rangeMap.put("down", range.getBottomRight().getRow() + index * betweenRowCount - 1);
							} else {
								int distanceTopBegin = range.getTopLeft().getRow() - beginRow;
								int distanceDownBegin = range.getBottomRight().getRow() - beginRow;
								rangeMap.put("left", range.getTopLeft().getColumn());
								rangeMap.put("top", distanceTopBegin + index * (betweenRowCount - betweenRowCountEach) - 1 + lastEachRow - 1);
								rangeMap.put("right", range.getBottomRight().getColumn());
								rangeMap.put("down", distanceDownBegin + index * (betweenRowCount - betweenRowCountEach) - 1 + lastEachRow - 1);
							}
							formatMergedList.add(rangeMap);
						}
						for (Integer rowPageBreak : rowPageBreakInEach) {
							formatrowPageBreakList.add(rowPageBreak + index * betweenRowCount - 1);
						}
						Map<Integer, Map> rowHeightMapMap = new HashMap<Integer, Map>();
						int tableFristRowCellMapMaxRow = 0;
						for (Map tableFristRowCellMap : tableFirstRows) {
							Cell tableFristRowCell = (Cell) tableFristRowCellMap.get("cell");
							if (!"$={end}".equals(tableFristRowCell.getContents()) && !"$={begin".equals(tableFristRowCell.getContents())) {
								Map eachCellMap = new HashMap();
								int eachCellMapRow = index * betweenRowCount + (Integer) tableFristRowCellMap.get("row") - 1;
								eachCellMap.put("cell", tableFristRowCell);
								eachCellMap.put("row", eachCellMapRow);
								eachCellMap.put("column", tableFristRowCell.getColumn());
								String cellContent = tableFristRowCell.getContents();
								if (eachCellMapRow > tableFristRowCellMapMaxRow) {
									tableFristRowCellMapMaxRow = eachCellMapRow;
								}
								if (rowHeightMapMap.get(eachCellMapRow) == null) {
									Map rowHeightMap = new HashMap();
									int rowHeight = readSheet.getRowHeight(tableFristRowCell.getRow());
									rowHeightMap.put("row", eachCellMapRow);
									rowHeightMap.put("height", rowHeight);
									rowHeightMapMap.put(eachCellMapRow, rowHeightMap);
								}
								if (!tableFristRowCell.getType().equals(CellType.EMPTY)) {
									Set<String> dataMapKeys = dataMap.keySet();
									for (String key : dataMapKeys) {
										Object dataContent = dataMap.get(key);
										String dataContentString = "";
										if (dataContent != null) {
											dataContentString = dataContent.toString();
										}
										cellContent = cellContent.replace("$={" + key + "}", dataContentString);
									}
								}
								eachCellMap.put("content", cellContent);
								eachValues.add(eachCellMap);
								lastEachRow = eachCellMapRow + 1;
							}
						}
						if (tableFristRowCellMapMaxRow != 0 && index == dataList.size() - 1) {
							rowHeightMapMap.remove(tableFristRowCellMapMaxRow);
						}
						Set<Integer> rowHeightMapSet = rowHeightMapMap.keySet();
						for (Integer rowHeight : rowHeightMapSet) {
							Map rowHeightMap = rowHeightMapMap.get(rowHeight);
							if (rowHeightMap != null) {
								rowHeightList.add(rowHeightMap);
							}
						}
					}
					if (i == beginList.size() - 1 && i == endList.size() - 1) {
						if (lastEachRow > 0) {
							lastEachRow = lastEachRow - 1;
							for (Range range : ranges) {
								if (range.getTopLeft().getRow() > endRow) {
									Map rangeMap = new HashMap();
									rangeMap.put("left", range.getTopLeft().getColumn());
									rangeMap.put("top", lastEachRow + range.getTopLeft().getRow() - endRow - 1);
									rangeMap.put("right", range.getBottomRight().getColumn());
									rangeMap.put("down", lastEachRow + range.getBottomRight().getRow() - endRow - 1);
									formatMergedList.add(rangeMap);
								}
							}
							for (int row = endRow + 1; row < rows; row++) {//columns
								for (int column = 0; column < columns; column++) {
									Cell oCell = readSheet.getCell(column, row);
									Map eachCellMap = new HashMap();
									eachCellMap.put("cell", oCell);
									eachCellMap.put("row", lastEachRow);
									eachCellMap.put("column", oCell.getColumn());
									eachCellMap.put("content", oCell.getContents());
									eachValues.add(eachCellMap);

									int rowHeight = readSheet.getRowHeight(row);
									Map rowHeightMap = new HashMap();
									rowHeightMap.put("row", lastEachRow);
									rowHeightMap.put("height", rowHeight);
									rowHeightList.add(rowHeightMap);
								}
								lastEachRow = lastEachRow + 1;
							}
						}
					}
				}
			}
		}
		Map<String, Object> formatMap = new HashMap<String, Object>();
		formatMap.put("formatMergedList", formatMergedList);
		formatMap.put("removeFormatMergedList", removeFormatMergedList);
		formatMap.put("formatrowPageBreakList", new ArrayList(formatrowPageBreakList));
		formatMap.put("removeRowList", removeRowList);
		formatMap.put("rowHeightList", rowHeightList);
		formatMap.put("lastEachRow", lastEachRow);
		writeExcel(outputStream, rwb, eachValues, itemList, itemValueMap, beginList, formatMap);
	}


	private void writeExcel(OutputStream outputStream, Workbook rwb, List<Map> eachValues, List<String> itemList, Map itemValueMap, List<Map> beginList, Map<String, Object> formatMap) throws Exception {
		WritableWorkbook writeBook = Workbook.createWorkbook(outputStream, rwb);
		rwb.close();
		WritableSheet firstSheet = writeBook.getSheet(0);
		List<Map> removeFormatMergedList = (List<Map>) formatMap.get("removeFormatMergedList");
		Range[] ranges = firstSheet.getMergedCells();
		Map<String, Range> rangeMap = new HashMap<String, Range>();
		for (Range range : ranges) {
			rangeMap.put(range.getTopLeft().getColumn() + "_" + range.getTopLeft().getRow() + "_" + range.getBottomRight().getColumn() + "_" + range.getBottomRight().getRow(), range);
		}
		List<Map> formatMergedList1 = (List<Map>) formatMap.get("formatMergedList");
		if (formatMergedList1 != null && formatMergedList1.size() > 0) {
			for (Map m : removeFormatMergedList) {
				int left = (Integer) m.get("left");
				int top = (Integer) m.get("top");
				int right = (Integer) m.get("right");
				int down = (Integer) m.get("down");
				Range range = rangeMap.get(left + "_" + top + "_" + right + "_" + down);
				if (range != null) {
					firstSheet.unmergeCells(range);
				}
			}
		}

		if (beginList.size() > 1) {
			Map m = beginList.get(0);
			if (m.get("error") == null || !(Boolean) m.get("error")) {
				firstSheet.removeRow((Integer) m.get("row"));
			}
		}
		for (Map m : eachValues) {
			Cell oCell = (Cell)m.get("cell");
			if(oCell != null){
				String content = "";
				if(m.get("content") != null){
					content = m.get("content").toString();
				}
				int column = (Integer)m.get("column");
				int row = (Integer)m.get("row");

				Label label = new Label(column, row, content);//判断item，获取对应的值  查询赋值
				if (oCell.getCellFeatures() != null) {
					WritableCellFeatures writableCellFeatures = new WritableCellFeatures(oCell.getCellFeatures());
					if (writableCellFeatures.hasDataValidation() && writableCellFeatures.getDVParser() == null) {
						List l = new ArrayList();
						l.add("");
						writableCellFeatures.setDataValidationList(l);
					}
					label.setCellFeatures(writableCellFeatures);
				}
				if (oCell.getCellFormat() != null) {
					WritableCellFormat writableCellFormat = new WritableCellFormat(oCell.getCellFormat());
					BorderLineStyle blsTop = oCell.getCellFormat().getBorderLine(Border.TOP);
					BorderLineStyle blsLeft = oCell.getCellFormat().getBorderLine(Border.LEFT);
					BorderLineStyle blsBottom = oCell.getCellFormat().getBorderLine(Border.BOTTOM);
					BorderLineStyle blsRight = oCell.getCellFormat().getBorderLine(Border.RIGHT);
					if (blsTop.getValue() == 1 && blsLeft.getValue() == 1 && blsBottom.getValue() == 0 && blsRight.getValue() == 0) {
						writableCellFormat.setBorder(Border.BOTTOM, blsTop);
						writableCellFormat.setBorder(Border.RIGHT, blsLeft);
					}
					label.setCellFormat(writableCellFormat);
				}
				firstSheet.addCell(label);
			}else{
				Label label = new Label((Integer)m.get("column"),(Integer)m.get("row"),null);//判断item，获取对应的值
				firstSheet.addCell(label);
			}
		}


		//循环执行结束之后，在标签里面赋值
		int rows = firstSheet.getRows();//获取工作表中的总行数
		int columns = firstSheet.getColumns();//获取工作表中的总列数
		for (int i = 0; i < rows; i++) {
			for (int j = 0; j < columns; j++) {
				Cell oCell = firstSheet.getCell(j, i);//需要注意的是这里的getCell方法的参数，第一个是指定第几列，第二个参数才是指定第几行
				if (!oCell.getType().equals(CellType.EMPTY)) {
					String contents = oCell.getContents();
					if (contents.indexOf("${") != -1) {
						if (contents.indexOf("}") != -1 && contents.indexOf("}") > contents.indexOf("${")) {
							//替换
							String content = contents;
							for (String s : itemList) {
								if (itemValueMap.get(s) != null) {
									content = content.replace("${" + s + "}", itemValueMap.get(s).toString());//对应itemList的值，从数据库查询   查询复制
								} else {
									content = content.replace("${" + s + "}", "");
								}
							}
							Label label = new Label(oCell.getColumn(), oCell.getRow(), content);//判断item。获取对应的值+
							if (oCell.getCellFeatures() != null) {

								WritableCellFeatures writableCellFeatures = new WritableCellFeatures(oCell.getCellFeatures());
								if (writableCellFeatures.hasDataValidation() && writableCellFeatures.getDVParser() == null) {
									List l = new ArrayList();
									l.add("");
									writableCellFeatures.setDataValidationList(l);
								}
								label.setCellFeatures(writableCellFeatures);
							}
							if (oCell.getCellFormat() != null) {
								WritableCellFormat writableCellFormat = new WritableCellFormat(oCell.getCellFormat());
								BorderLineStyle blsTop = oCell.getCellFormat().getBorderLine(Border.TOP);
								BorderLineStyle blsLeft = oCell.getCellFormat().getBorderLine(Border.LEFT);
								BorderLineStyle blsBottom = oCell.getCellFormat().getBorderLine(Border.BOTTOM);
								BorderLineStyle blsRight = oCell.getCellFormat().getBorderLine(Border.RIGHT);
								if (blsTop.getValue() == 1 && blsLeft.getValue() == 1 && blsBottom.getValue() == 0 && blsRight.getValue() == 0) {
									writableCellFormat.setBorder(Border.BOTTOM, blsTop);
									writableCellFormat.setBorder(Border.RIGHT, blsLeft);
								}
								label.setCellFormat(writableCellFormat);
							}
							firstSheet.addCell(label);
						}
					}
				} else {
					Label label = new Label(oCell.getColumn(), oCell.getRow(), null);//判断item。获取对应的值
					if (oCell.getCellFeatures() != null) {
						WritableCellFeatures writableCellFeatures = new WritableCellFeatures(oCell.getCellFeatures());
						if (writableCellFeatures.hasDataValidation() && writableCellFeatures.getDVParser() == null) {
							List l = new ArrayList();
							l.add("");
							writableCellFeatures.setDataValidationList(l);
						}
						label.setCellFeatures(writableCellFeatures);
					}
					if (oCell.getCellFormat() != null) {
						WritableCellFormat writableCellFormat = new WritableCellFormat(oCell.getCellFormat());
						label.setCellFormat(writableCellFormat);
					}
					firstSheet.addCell(label);
				}
			}
		}

		List<Map> formatMergedList = (List<Map>) formatMap.get("formatMergedList");
		if (formatMergedList != null) {
			for (Map mergedMap : formatMergedList) {
				int left = (Integer) mergedMap.get("left");
				int top = (Integer) mergedMap.get("top");
				int right = (Integer) mergedMap.get("right");
				int down = (Integer) mergedMap.get("down");
				firstSheet.mergeCells(left, top, right, down);
			}
		}
		List<Integer> formatrowPageBreakList = (List<Integer>) formatMap.get("formatrowPageBreakList");
		for (Integer rowPageBreak : formatrowPageBreakList) {
			firstSheet.addRowPageBreak(rowPageBreak);
		}

		List<Map> rowHeightList = (List<Map>) formatMap.get("rowHeightList");
		for (Map m : rowHeightList) {
			int row = (Integer) m.get("row");
			int rowHeight = (Integer) m.get("height");
			firstSheet.setRowView(row, rowHeight);
		}

		int rowMax = firstSheet.getRows();
		int colMax = firstSheet.getColumns();
		int lastEachRow = (Integer) formatMap.get("lastEachRow");
		for (int i = 0; i < rowMax; i++) {
			for (int j = 0; j < colMax; j++) {
				if (i >= lastEachRow && lastEachRow != 0) {
					firstSheet.removeRow(lastEachRow);
					continue;
				}
				Cell cell = firstSheet.getCell(j, i);
				if (cell.getContents() != null) {
					if (cell.getContents().indexOf("$={end}") != -1 || cell.getContents().indexOf("$={begin:") != -1) {
						firstSheet.removeRow(cell.getRow());
					}
				}
			}
		}
		List<Map> removeRowList = (List<Map>) formatMap.get("removeRowList");
		for (Map m : removeRowList) {
			if (m.get("error") == null || !(Boolean) m.get("error")) {
				firstSheet.removeRow((Integer) m.get("row"));
			}
		}
		// 4、打开流，开始写文件
		writeBook.write();
		// 5、关闭流
		writeBook.close();
	}

	public static void main(String[] args) {

	}
}