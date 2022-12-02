package com.laipan.factory;

import com.laipan.utils.CommonUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;

/**
 * <p><p/>
 *
 * @author laipan
 * @date 2022/12/02,10:19
 * @since v0.1
 */
public class ExcelFactory {

    private final static String EXCEL_HIGHER_VERSION = "xlsx";
    private final static String EXCEL_LOWER_VERSION  = "xls";

    private static Workbook wb;

    public static Sheet createSheet(String fileName, InputStream is, int sheetNum) throws IOException {
        String fileType = CommonUtil.getFileType(fileName);
        // excel 文件内容压缩率限制，-1.0d 表示无此限制，可能存在解压之后内容过多导致 OOM 的风险
        ZipSecureFile.setMinInflateRatio(-1.0d);
        if (EXCEL_HIGHER_VERSION.equalsIgnoreCase(fileType)) {
            wb = new XSSFWorkbook(is);
        } else {
            wb = new HSSFWorkbook(is);
        }

        return wb.getSheetAt(sheetNum);
    }
}
