package com.laipan.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * <p><p/>
 *
 * @author laipan
 * @date 2022/12/05,17:41
 * @since v0.1
 */
public class ExcelUtils {

    /**
     * 获取单元格各类型值，返回字符串类型
     *
     * @param cell 单元格
     * @return 空串当作 null 处理
     */
    public static String getCellValueByCell(Cell cell) {
        //判断是否为 null 或空串，空串也当作 null
        if (cell == null || cell.toString().trim().equals("")) {
            return null;
        }
        String   cellValue = "";
        CellType cellType  = cell.getCellType();
        switch (cellType) {
            case NUMERIC: // 数字
                short format = cell.getCellStyle().getDataFormat();
                if (DateUtil.isCellDateFormatted(cell)) {
                    SimpleDateFormat sdf = null;
                    if (format == 20 || format == 32) {
                        sdf = new SimpleDateFormat("HH:mm");
                    } else if (format == 14 || format == 31 || format == 57 || format == 58) {
                        // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
                        sdf = new SimpleDateFormat("yyyy-MM-dd");
                        double value = cell.getNumericCellValue();
                        Date date = org.apache.poi.ss.usermodel.DateUtil
                                .getJavaDate(value);
                        cellValue = sdf.format(date);
                    } else {// 日期
                        sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    }
                    try {
                        cellValue = sdf.format(cell.getDateCellValue());// 日期
                    } catch (Exception e) {
                        try {
                            throw new Exception("exception on get date data !".concat(e.toString()));
                        } catch (Exception e1) {
                            e1.printStackTrace();
                        }
                    } finally {
                        sdf = null;
                    }
                } else {
                    BigDecimal bd = new BigDecimal(cell.getNumericCellValue());
                    cellValue = bd.toPlainString();// 数值 这种用BigDecimal包装再获取plainString，可以防止获取到科学计数值
                }
                break;
            case STRING: // 字符串
                cellValue = cell.getStringCellValue();
                break;
            case BOOLEAN: // Boolean
                cellValue = cell.getBooleanCellValue() + "";
                break;
            case FORMULA: // 公式
                // 如果直接获取公式类型的数据，则获取直接就是公式本身，如：1+1，""员工"" 等
                if (cell.getCellFormula().startsWith("EOMONTH")) {
                    // Excel 的 EOMONTH 函数在计算日期不会返回 string 类型的日期数据，而是返回日期序列号，1900-01-01 的序列号为 1，按日递增。
                    SimpleDateFormat sdf = null;
                    try {
                        Date date = cell.getDateCellValue();
                        sdf = new SimpleDateFormat("yyyy-MM-dd");
                        cellValue = sdf.format(date);
                    } catch (Exception e) {
                        try {
                            throw new Exception("exception on get date data within 【EOMONTH】 function !".concat(e.toString()));
                        } catch (Exception e1) {
                            e1.printStackTrace();
                        }
                    } finally {
                        sdf = null;
                    }
                } else {
                    cellValue = ((XSSFCell) cell).getRawValue();
                }
                break;
            case BLANK: // 空值
                cellValue = "";
                break;
            case ERROR: // 故障
                cellValue = "ERROR VALUE";
                break;
            default:
                cellValue = "UNKNOW VALUE";
                break;
        }
        return cellValue;
    }
}
