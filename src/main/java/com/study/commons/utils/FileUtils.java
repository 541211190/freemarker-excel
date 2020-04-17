
package com.study.commons.utils;

import java.io.File;

/**
 * @project cne-power-commons-util
 * @description: 操作文件和文件夹工具类
 * @author 大脑补丁
 * @create 2020-04-05 11:22
 */
public class FileUtils {

    /**
     * 创建该路径下，文件夹和文件
     * 
     * @param filePath
     * @return
     * @author 大脑补丁 on 2020-04-05 11:24
     */
    public static File createFile(String filePath) {
        File file = new File(filePath);
        if (!file.exists()) {
            if (!file.getParentFile().exists()) {
                // 若父级目录不存在，先创建父级目录
                file.getParentFile().mkdirs();
            }
            try {
                // 创建文件
                file.createNewFile();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return file;
    }
}
