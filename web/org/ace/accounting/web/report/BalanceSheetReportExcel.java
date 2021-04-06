package org.ace.accounting.web.report;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.ace.accounting.common.Utils;
import org.ace.accounting.common.utils.BusinessUtil;
import org.ace.accounting.report.balancesheet.BalanceSheetCriteria;
import org.ace.accounting.report.balancesheet.BalanceSheetDTO;
import org.ace.accounting.web.common.ExcelUtils;
import org.ace.java.component.ErrorCode;
import org.ace.java.component.SystemException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class BalanceSheetReportExcel {

	private XSSFWorkbook wb;
	// private FormulaEvaluator evaluator ;

	public BalanceSheetReportExcel() {
		load();
	}

	private void load() {
		try {
			InputStream inp = this.getClass().getResourceAsStream("/report-template/Report.xlsx");
			wb = new XSSFWorkbook(inp);
		} catch (IOException e) {
			throw new SystemException(ErrorCode.SYSTEM_ERROR, "Failed to load Report.xlsx tempalte", e);
		}
	}

	public Map<String, List<BalanceSheetDTO>> separateByPaymentChannel(List<BalanceSheetDTO> orderList) {
		Map<String, List<BalanceSheetDTO>> map = new LinkedHashMap<String, List<BalanceSheetDTO>>();
		if (orderList != null) {
			for (BalanceSheetDTO report : orderList) {
				if (map.containsKey(report.getAcCode())) {
					map.get(report.getAcCode()).add(report);
				} else {
					List<BalanceSheetDTO> list = new ArrayList<BalanceSheetDTO>();
					list.add(report);
					map.put(report.getAcCode(), list);
				}
			}
		}
		return map;

	}

	public void generate(OutputStream op, List<BalanceSheetDTO> orderList, BalanceSheetCriteria criteria) {
		try {
			Sheet sheet = wb.getSheet("Report");
			String type = "";

			if (criteria.isMonth()) {
				type = "Monthly ";
			} else {
				type = "Date Range ";
			}
			if (criteria.getReportType().equalsIgnoreCase("B")) {
				wb.setSheetName(wb.getSheetIndex("Report"), "BalanceSheet");
				type = type.concat("Balance Sheet ");
			} else {
				wb.setSheetName(wb.getSheetIndex("Report"), "Profit&Loss");
				type = type.concat("Profit And Loss ");
			}

			XSSFCellStyle defaultCellStyle = ExcelUtils.getDefaultCellStyle(wb);
			XSSFCellStyle textCellStyle = ExcelUtils.getTextCellStyle(wb);
			XSSFCellStyle dateCellStyle = ExcelUtils.getDateCellStyle(wb);
			XSSFCellStyle currencyCellStyle = ExcelUtils.getCurrencyCellStyle(wb);
			XSSFCellStyle textAlignRightStyle = ExcelUtils.getTextAlignRightStyle(wb);
			XSSFCellStyle centerCellStyle = ExcelUtils.getAlignCenterStyle(wb);
			XSSFCellStyle textAlignCenterStyle = ExcelUtils.getAlignCenterStyle(wb);

			Row row;
			Cell cell;
			row = sheet.getRow(0);
			cell = row.getCell(0);
			String branch = "";
			String currency = "";

			if (criteria.getBranch() == null) {
				branch = "All Branches";
			} else {
				branch = criteria.getBranch().getName();
			}

			if (criteria.getCurrency() == null) {
				currency = "All Currencies";
			} else {
				currency = criteria.getCurrency().getCurrencyCode();
			}
			if (criteria.isHomeConverted() && criteria.getCurrency() != null) {
				currency = currency + " By Home Currency Converted";
			}

			row = sheet.getRow(1);
			cell = row.createCell(7);
			String dateValue = "";
			if (criteria.isMonth()) {
				dateValue = type + "as at " + Utils.getDateFormatString(BusinessUtil.getBudgetStartDate()) + " To " + Utils.getDateFormatString(BusinessUtil.getBudgetEndDate());
			} else {
				dateValue = type + "as at " + Utils.getDateFormatString(criteria.getStartDate()) + " To " + Utils.getDateFormatString(criteria.getEndDate());
			}
			cell.setCellValue(dateValue);
			cell.setCellStyle(textAlignCenterStyle);

			row = sheet.getRow(2);
			cell = row.getCell(1);
			cell.setCellValue(branch);
			// cell.setCellStyle(defaultCellStyle);

			row = sheet.getRow(3);
			cell = row.getCell(1);
			cell.setCellValue(currency);
			// cell.setCellStyle(defaultCellStyle);

			// cell = row.createCell(13);
			// cell.setCellValue("Date : ");
			// cell.setCellStyle(defaultCellStyle);

			cell = row.createCell(14);
			cell.setCellValue(Utils.getDateFormatString(new Date()));
			// cell.setCellStyle(dateCellStyle);

			int i = 4;
			int index = 0;

			for (BalanceSheetDTO report : orderList) {

				i = i + 1;
				index = index + 1;

				row = sheet.createRow(i);
				cell = row.createCell(0);
				cell.setCellValue(index);
				cell.setCellStyle(textCellStyle);

				cell = row.createCell(1);
				cell.setCellValue(report.getAcCode());
				cell.setCellStyle(textCellStyle);

				cell = row.createCell(2);
				cell.setCellValue(report.getAcName());
				cell.setCellStyle(textCellStyle);

				cell = row.createCell(3);
				if (null != report.getM1()) {
					cell.setCellValue(Double.valueOf(report.getM1().toString()));
					cell.setCellStyle(currencyCellStyle);
				} else {
					cell.setCellValue(0);
					cell.setCellStyle(currencyCellStyle);
				}
				cell = row.createCell(4);
				if (null != report.getM2()) {
					cell.setCellValue(Double.valueOf(report.getM2().toString()));
					cell.setCellStyle(currencyCellStyle);
				} else {
					cell.setCellValue(0);
					cell.setCellStyle(currencyCellStyle);
				}

				cell = row.createCell(5);
				if (null != report.getM3()) {
					cell.setCellValue(Double.valueOf(report.getM3().toString()));
					cell.setCellStyle(currencyCellStyle);
				} else {
					cell.setCellValue(0);
					cell.setCellStyle(currencyCellStyle);
				}

				cell = row.createCell(6);
				if (null != report.getM4()) {
					cell.setCellValue(Double.valueOf(report.getM4().toString()));
					cell.setCellStyle(currencyCellStyle);
				} else {
					cell.setCellValue(0);
					cell.setCellStyle(currencyCellStyle);
				}

				cell = row.createCell(7);
				if (null != report.getM5()) {
					cell.setCellValue(Double.valueOf(report.getM5().toString()));
					cell.setCellStyle(currencyCellStyle);
				} else {
					cell.setCellValue(0);
					cell.setCellStyle(currencyCellStyle);
				}

				cell = row.createCell(8);
				if (null != report.getM6()) {
					cell.setCellValue(Double.valueOf(report.getM6().toString()));
					cell.setCellStyle(currencyCellStyle);
				} else {
					cell.setCellValue(0);
					cell.setCellStyle(currencyCellStyle);
				}

				cell = row.createCell(9);
				if (null != report.getM7()) {
					cell.setCellValue(Double.valueOf(report.getM7().toString()));
					cell.setCellStyle(currencyCellStyle);
				} else {
					cell.setCellValue(0);
					cell.setCellStyle(currencyCellStyle);
				}
				cell = row.createCell(10);
				if (null != report.getM8()) {
					cell.setCellValue(Double.valueOf(report.getM8().toString()));
					cell.setCellStyle(currencyCellStyle);
				} else {
					cell.setCellValue(0);
					cell.setCellStyle(currencyCellStyle);
				}

				cell = row.createCell(11);
				if (null != report.getM9()) {
					cell.setCellValue(Double.valueOf(report.getM9().toString()));
					cell.setCellStyle(currencyCellStyle);
				} else {
					cell.setCellValue(0);
					cell.setCellStyle(currencyCellStyle);
				}
				cell = row.createCell(12);
				if (null != report.getM10()) {
					cell.setCellValue(Double.valueOf(report.getM10().toString()));
					cell.setCellStyle(currencyCellStyle);
				} else {
					cell.setCellValue(0);
					cell.setCellStyle(currencyCellStyle);
				}

				cell = row.createCell(13);
				if (null != report.getM11()) {
					cell.setCellValue(Double.valueOf(report.getM11().toString()));
					cell.setCellStyle(currencyCellStyle);
				} else {
					cell.setCellValue(0);
					cell.setCellStyle(currencyCellStyle);
				}

				cell = row.createCell(14);
				if (null != report.getM12()) {
					cell.setCellValue(Double.valueOf(report.getM12().toString()));
					cell.setCellStyle(currencyCellStyle);
				} else {
					cell.setCellValue(0);
					cell.setCellStyle(currencyCellStyle);
				}
				// paymentChannel = report.getPaymentChannel().toString();
			}
			i = i + 1;
			sheet.addMergedRegion(new CellRangeAddress(i, i, 0, 14));
			row = sheet.createRow(i);
			cell = row.createCell(0);

			cell = row.createCell(0);
			cell.setCellStyle(defaultCellStyle);
			cell.setCellValue("-");

			cell = row.createCell(1);
			cell.setCellStyle(defaultCellStyle);
			cell.setCellValue("-");

			cell = row.createCell(2);
			cell.setCellStyle(defaultCellStyle);
			cell.setCellValue("-");

			cell = row.createCell(3);
			cell.setCellStyle(defaultCellStyle);
			cell.setCellValue("-");

			cell = row.createCell(4);
			cell.setCellStyle(defaultCellStyle);
			cell.setCellValue("-");
			cell = row.createCell(5);
			cell.setCellStyle(defaultCellStyle);
			cell.setCellValue("-");

			cell = row.createCell(6);
			cell.setCellStyle(defaultCellStyle);
			cell.setCellValue("-");

			cell = row.createCell(7);
			cell.setCellStyle(defaultCellStyle);
			cell.setCellValue("-");

			cell = row.createCell(8);
			cell.setCellStyle(defaultCellStyle);
			cell.setCellValue("-");

			cell = row.createCell(9);
			cell.setCellStyle(defaultCellStyle);
			cell.setCellValue("-");

			cell = row.createCell(10);
			cell.setCellStyle(defaultCellStyle);
			cell.setCellValue("-");

			cell = row.createCell(11);
			cell.setCellStyle(defaultCellStyle);
			cell.setCellValue("-");

			cell = row.createCell(12);
			cell.setCellStyle(defaultCellStyle);
			cell.setCellValue("-");

			cell = row.createCell(13);
			cell.setCellStyle(defaultCellStyle);
			cell.setCellValue("-");

			cell = row.createCell(14);
			cell.setCellStyle(defaultCellStyle);
			cell.setCellValue("-");

			wb.setPrintArea(0, 0, 14, 0, i);
			wb.write(op);
			op.flush();
			op.close();
			// }
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
