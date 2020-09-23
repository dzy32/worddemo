package com.example.worddemo.util;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

/**
 * @author ys
 * @date 2020/9/21 11:52
 */
public class OtherUtil {
    /**
     * 指定路径读取文件
     * (未测试大文件读取、推荐使用进行小文件读取)
     * @return
     * @throws IOException
     */
    public static String getFileByPath(String path) throws IOException {
        return new String(Files.readAllBytes(Paths.get(path)));
    }

}
