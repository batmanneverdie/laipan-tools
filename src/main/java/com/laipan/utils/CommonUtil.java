package com.laipan.utils;

/**
 * <p><p/>
 *
 * @author laipan
 * @date 2022/03/11,16:50
 * @since v0.1
 */
public class CommonUtil {


    /**
     * 获取文件类型
     *
     * @param fileName
     * @return
     */
    public static String getFileType(String fileName) {
        return fileName.substring(fileName.lastIndexOf(".") + 1).toLowerCase();
    }
}