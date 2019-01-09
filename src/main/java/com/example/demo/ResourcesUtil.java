package com.example.demo;

/**
 * @author lan_bing
 * @date 2019-01-09 12:19
 */
/**
 * 描述：获取项目中resources路径
 *
 * @author ssl
 * @create 2018/08/23 11:13
 */
public class ResourcesUtil {
    public static String getClasspath() {
        String path = Thread.currentThread().getContextClassLoader().getResource("").toString();
        if (path.startsWith("file:")) {
            path = path.substring(5);
        }
        return path;
    }

    public static String getClasspathFile(String path) {
        String file = "";
        file = Thread.currentThread().getContextClassLoader().getResource(path).getFile();
        return file;
    }

    public static void main(String[] args) {
        System.out.println(getClasspath());
        System.out.println(getClasspathFile("yc1.pdf"));
    }
}