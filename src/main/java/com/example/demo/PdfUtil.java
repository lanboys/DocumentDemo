package com.example.demo;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.itextpdf.forms.PdfAcroForm;
import com.itextpdf.forms.fields.PdfTextFormField;
import com.itextpdf.io.font.FontProgram;
import com.itextpdf.io.font.FontProgramFactory;
import com.itextpdf.io.font.PdfEncodings;
import com.itextpdf.kernel.colors.ColorConstants;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.Rectangle;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.PdfWriter;

import org.springframework.util.ResourceUtils;

import java.io.File;
import java.io.IOException;

/**
 * 描述：操作pdf的工具类《依赖itext7 7.1.2》
 *
 * @author ssl
 * @create 2018/08/23 10:07
 * https://gitee.com/shisenlin/tool.git
 * https://gitee.com/shisenlin/tool/blob/master/src/main/java/com/ssl/tools/pdf/PdfUtil.java#
 * https://blog.csdn.net/ZuoYanYouYan/article/details/81974656
 */
public class PdfUtil {

    /**
     * 在pdf上创建表单域
     *
     * @param pdf
     * @param savePath
     * @param fields：  [{
     *                 "page": "1",
     *                 "positionX": "x轴的距离",
     *                 "positionY": "y轴的距离",
     *                 "width": "长",
     *                 "height": "宽",
     *                 "font":"字体，ttf格式",
     *                 "fontSize": "字体大小",
     *                 "name":"表单域的名称",
     *                 "value":"表单域的值，可选"
     *                 },
     *                 {
     *                 "page": "1",
     *                 "positionX": "x轴的距离",
     *                 "positionY": "y轴的距离",
     *                 "width": "长",
     *                 "height": "宽",
     *                 "font":"字体，ttf格式",
     *                 "fontSize": "字体大小",
     *                 "name":"表单域的名称",
     *                 "value":"表单域的值，可选"
     *                 }
     *                 ]
     * @throws IOException
     */
    public static void createText(File pdf, String savePath, JSONArray fields) throws IOException {
        // 编辑后的文件
        PdfWriter pdfWriter = new PdfWriter(savePath);
        PdfDocument pdfDocument = new PdfDocument(new PdfReader(pdf), pdfWriter);
        PdfAcroForm pdfAcroForm = PdfAcroForm.getAcroForm(pdfDocument, true);
        //font
        String ttfPath = ResourceUtils.getURL(ResourceUtils.CLASSPATH_URL_PREFIX.concat("font")).getPath();
        FontProgram fontProgram;
        PdfFont font;
        fontProgram = FontProgramFactory.createFont("G:\\intelijIdea_workplace\\tool\\src\\main" +
                "\\resources\\font\\STKAITI.TTF");
        font = PdfFontFactory.createFont(fontProgram, PdfEncodings.IDENTITY_H, true);

        PdfTextFormField pdfTextFormField;
        JSONObject field;
        Rectangle rectangle;
        for (int i = 0; i < fields.size(); i++) {
            field = fields.getJSONObject(i);
            // 读取ttf字体文件
            fontProgram = FontProgramFactory.createFont(ttfPath.concat(File.separator).concat(field.getString
                    ("font")));
            // 编码使用 PdfEncodings.IDENTITY_H
            font = PdfFontFactory.createFont(fontProgram, PdfEncodings.IDENTITY_H, true);
            // 设置表单域的位置
            rectangle = new Rectangle(field.getFloatValue("positionX"), field.getFloatValue("positionY"), field
                    .getFloatValue("width"), field.getFloatValue("height"));
            pdfTextFormField = PdfTextFormField.createText(pdfDocument, rectangle, field.getString("name"), field
                    .getString("value"));
            pdfTextFormField.setBorderWidth(0).setReadOnly(true).setColor(ColorConstants.BLACK).setFontAndSize(font,
                    field.getIntValue("fontSize"));
            /** 将表单域加入pdf的指定页中 */
            pdfAcroForm.addField(pdfTextFormField, pdfDocument.getPage(field.getIntValue("page")));
        }
        // 将表单域中的value嵌入到pdf文件中
        pdfAcroForm.flattenFields();
        pdfDocument.close();
        pdfWriter.close();
    }

    public static void main(String[] args) throws IOException {
        //表单域
        JSONArray jsonArray = JSONArray.parseArray("[{\"fontSize\":\"18\",\"height\":20,\"name\":\"field1\"," +
                "\"page\":1," +
                "\"positionX\":0,\"positionY\":10,\"value\":\"这是华文幼圆的中文字\",\"font\":\"SIMYOU.TTF\",\"width\":200}," +
                "{\"fontSize\":\"18\"," +
                "\"height\":20,\"name\":\"field2\",\"page\":2,\"positionX\":0,\"positionY\":820," +
                "\"value\":\"这是楷体的中文字\"," +
                "\"font\":\"STKAITI.TTF\"," +
                "\"width\":200}]");
        //源文件
        File pdf = new File(ResourcesUtil.getClasspathFile("stdlib-3.4.pdf"));
        // 保存路径
        String savePath = "G:\\intelijIdea_workplace\\tool\\src\\main\\resources\\stdlib-3.4-update.pdf";
        createText(pdf, savePath, jsonArray);
    }
}
