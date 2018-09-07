package main;

import java.awt.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

/**
 * 工具类
 */
class Utils {
    static String getCurrentTimeStr() {
        LocalDateTime now = LocalDateTime.now();
        return now.format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss:SSS"));
    }

    static String getCurrentTimeStr2() {
        LocalDateTime now = LocalDateTime.now();
        return now.format(DateTimeFormatter.ofPattern("yyyyMMddHHmmss"));
    }

    static Font myFont(int fontSize) {
        return new Font("宋体", Font.PLAIN, fontSize);
    }
}
