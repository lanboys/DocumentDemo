package com.example.demo;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.AcroFields;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;

import java.io.FileOutputStream;
import java.io.IOException;

public class GeneratePDFUtil {

    //测试
    public static void main(String[] args) throws IOException, DocumentException {

        String templatePath = "D:\\workspace\\IDEA_workspace\\PdfDemo\\src\\main\\resources\\temp.pdf";
        // 生成的新文件路径
        String newPDFPath = "D:\\workspace\\IDEA_workspace\\PdfDemo\\src\\main\\resources\\file.pdf";

        manipulatePdf(templatePath, newPDFPath);
    }

    public static void manipulatePdf(String src, String dest) throws DocumentException, IOException {
        PdfReader reader = new PdfReader(src);
        PdfStamper stamper = new PdfStamper(reader, new FileOutputStream(dest));
        AcroFields fields = stamper.getAcroFields();

        // 这么设置字体 不起作用
        //BaseFont baseFont = BaseFont.createFont("simsun.ttc,0", "Identity-H", BaseFont.NOT_EMBEDDED);
        //fields.setFieldProperty("name", "textfont", baseFont, null);
        fields.setFieldProperty("name", "textcolor", BaseColor.BLUE, null);
        fields.setFieldProperty("name", "bordercolor", BaseColor.RED, null);
        fields.setFieldProperty("name", "fontsize", 14, null);

        // 给表单添加中文字体 这里采用系统字体。不设置的话，中文可能无法显示
        //https://bbs.csdn.net/topics/390924324
        //https://blog.csdn.net/qq_35893120/article/details/82786066

        //设置字体
        //BaseFont bf = BaseFont.createFont("simsun.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        //fields.addSubstitutionFont(bf);

        fields.setField("name", "中文杠杠滴郭德纲的");
        fields.setField("idNo", "中文杠杠滴郭德纲的");

        // 不设置字体 设置为 true 不显示
        // 不设置字体 设置为 false 点文本域显示 能编辑

        // 设置字体 设置为 true  显示
        // 设置字体 设置为 false 显示 但是能编辑
        stamper.setFormFlattening(true);// 如果为false那么生成的PDF文件还能编辑，一定要设为true

        stamper.close();
        reader.close();
    }
}
