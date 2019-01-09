package com.example.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * @author lan_bing
 * @date 2019-01-07 10:09
 */
public class FileUtils {

    /**
     * 文件转base64字符串
     *
     * @param file
     * @return
     */
    public static byte[] fileToBytes(File file) {
        InputStream in = null;
        byte[] bytes = null;
        try {
            in = new FileInputStream(file);
            bytes = new byte[in.available()];
            in.read(bytes);
            System.out.println("fileToBytes() length: " + bytes.length);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (in != null) {
                    in.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return bytes;
    }

    public static byte[] stringToBytes(String str) {
        byte[] bytes = str.getBytes();
        System.out.println("stringToBytes() length: " + bytes.length);
        return bytes;
    }

    public static void cover(File file) throws IOException {
        InputStream input = new FileInputStream(file);

        StringBuffer out = new StringBuffer();
        byte[] b = new byte[4096];
        for (int n; (n = input.read(b)) != -1; ) {
            out.append(new String(b, 0, n));
        }

        System.out.println("cover() : " + out.toString());
    }
}
