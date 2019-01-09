package com.example.demo;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.AcroFields;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfCopy;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

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

    public static String pdfOut(String templatePath, File destPath, Map<String, Object> o) {
        try {
            FileOutputStream out = new FileOutputStream(destPath);// 输出流
            PdfReader reader = new PdfReader(templatePath);// 读取pdf模板
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            PdfStamper stamper = new PdfStamper(reader, bos);

            AcroFields form = stamper.getAcroFields();
            //Font FontChinese = new Font(bf, 5, Font.NORMAL);

            //原文：https://blog.csdn.net/sand_clock/article/details/70227077
            //String fontName = "";
            //String os = System.getProperties().getProperty("os.name");
            //if (os.startsWith("win") || os.startsWith("Win")) {
            //    //fontName = DocumentUtil.class.getResource("/fonts").getPath() + "simsun.ttc,0";
            //    fontName = "C:/Windows/Fonts/simsun.ttc,0";
            //    BaseFont bf = BaseFont.createFont(fontName, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            //    form.addSubstitutionFont(bf);
            //} else {
            //    fontName = "/usr/share/fonts/lyx/simsun.ttc";
            //    BaseFont bf = BaseFont.createFont(fontName, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            //    form.addSubstitutionFont(bf);
            //}

            form.setFieldProperty("text", "textcolor", BaseColor.BLUE, null);
            form.setFieldProperty("text", "bordercolor", BaseColor.RED, null);
            form.setFieldProperty("text", "fontsize", 34, null);

            //文字类的内容处理
            Map<String, String> datemap = (Map<String, String>) o.get("dateMap");
            if (datemap != null) {
                for (String key : datemap.keySet()) {
                    String value = datemap.get(key);
                    form.setField(key, value);
                }
            }
            //图片类的内容处理
            Map<String, String> imgmap = (Map<String, String>) o.get("imgMap");
            if (imgmap != null) {
                for (String key : imgmap.keySet()) {
                    String value = imgmap.get(key);
                    int pageNo = form.getFieldPositions(key).get(0).page;
                    Rectangle signRect = form.getFieldPositions(key).get(0).position;
                    float x = signRect.getLeft();
                    float y = signRect.getBottom();
                    //根据路径读取图片
                    Image image = Image.getInstance(value);
                    //获取图片页面
                    PdfContentByte under = stamper.getOverContent(pageNo);
                    //图片大小自适应
                    image.scaleToFit(signRect.getWidth(), signRect.getHeight());
                    //添加图片
                    image.setAbsolutePosition(x, y);
                    under.addImage(image);
                }
            }
            stamper.setFormFlattening(false);// 如果为false，生成的PDF文件可以编辑，如果为true，生成的PDF文件不可以编辑
            stamper.close();

            //https://zhuchengzzcc.iteye.com/blog/1603671
            Document doc = new Document();
            PdfCopy copy = new PdfCopy(doc, out);
            doc.open();
            PdfReader reader1 = new PdfReader(bos.toByteArray());
            int numberOfPages = reader1.getNumberOfPages();
            for (int i = 0; i < numberOfPages; i++) {
                PdfImportedPage importPage = copy.getImportedPage(reader1, i + 1);
                copy.addPage(importPage);
            }
            doc.close();
            out.close();
            bos.close();
        } catch (Exception e) {
            //logger.info("利用模板生成pdf异常: ", e);
            return null;
        }
        return destPath.getAbsolutePath();
    }

    // 利用模板生成pdf
    public static void interviewReportPDF(Map<String, String> map) {

        // 模板路径
        //String templatePath = UtilPath.getWEB_INF() + "pdf/面试报告模板.pdf";
        String templatePath = "D:\\workspace\\IDEA_workspace\\MSJApiDemo\\src\\main\\resources\\bbn_wrapper.pdf";
        // 生成的新文件路径
        String newPDFPath = "D:\\workspace\\IDEA_workspace\\MSJApiDemo\\src\\main\\resources\\bbnaqqqq.pdf";
        PdfReader reader;
        FileOutputStream out;
        ByteArrayOutputStream bos;
        PdfStamper stamper;
        try {
            out = new FileOutputStream(newPDFPath);// 输出流
            reader = new PdfReader(templatePath);// 读取pdf模板
            bos = new ByteArrayOutputStream();
            stamper = new PdfStamper(reader, bos);
            AcroFields form = stamper.getAcroFields();

            // 给表单添加中文字体 这里采用系统字体。不设置的话，中文可能无法显示
            //https://bbs.csdn.net/topics/390924324
            //https://blog.csdn.net/qq_35893120/article/details/82786066
            //BaseFont bf = BaseFont.createFont(UtilPath.getRootPath() + "fonts/simsun.ttc,0", BaseFont.IDENTITY_H,
            //        BaseFont.EMBEDDED);
            //form.addSubstitutionFont(bf);
            //BaseFont baseFont = BaseFont.createFont("D:\\workspace\\IDEA_workspace\\MSJApiDemo\\src\\main\\resources\\simsun.ttc,0", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            //form.addSubstitutionFont(baseFont);

            form.setFieldProperty("name", "textcolor", BaseColor.BLUE, null);
            form.setFieldProperty("name", "bordercolor", BaseColor.RED, null);
            form.setFieldProperty("name", "fontsize", 14, null);

            //遍历map装入数据
            for (Map.Entry<String, String> entry : map.entrySet()) {
                form.setField(entry.getKey(), entry.getValue());
                System.out.println("插入PDF数据---->  key= " + entry.getKey() + " and value= " + entry.getValue());
            }

            stamper.setFormFlattening(false);// 如果为false那么生成的PDF文件还能编辑，一定要设为true
            stamper.close();
            reader.close();

            // 拷贝合并 操作 可以直接返回流
            // https://zhuchengzzcc.iteye.com/blog/1603671
            Document doc = new Document();
            PdfCopy copy = new PdfCopy(doc, out);
            doc.open();
            PdfReader reader1 = new PdfReader(bos.toByteArray());
            int numberOfPages = reader1.getNumberOfPages();
            for (int i = 2; i < numberOfPages + 1; i++) {
                PdfImportedPage importPage = copy.getImportedPage(reader1, i);
                copy.addPage(importPage);
            }
            doc.close();
            out.close();

            //Document document1 = new Document();
            //String destPath1 = "E:\\workspace\\IDEA_workspace\\MSJApiDemo\\src\\main\\resources\\dd.pdf";
            //
            //FileOutputStream outputStream = new FileOutputStream(destPath1);
            //PdfWriter.getInstance(document1, outputStream);
            //
            //document1.open();
            //document1.add(new Paragraph("Hello World"));
            //document1.close();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (DocumentException e) {
            e.printStackTrace();
        }
    }

    /**
     * https://blog.csdn.net/weixin_41187876/article/details/79156969
     * https://blog.csdn.net/top__one/article/details/65442390
     */
    public static void main1(String[] args) throws IOException, DocumentException {
        Map<String, String> map = new HashMap<>();
        map.put("name", "蓝兵");
        map.put("idNo", "哈哈哈");

        System.out.println(map);
        interviewReportPDF(map);
    }

    static String x = "个人授权协议\n" +
            "\n" +
            "个人征信授权协议（以下简称本协议）是由北京想就拿信息技术有限公司（以下简称本公司/我们）和您签订。如您点击选择同意本协议将视为授权本公司向第三方支付／征信／金融机构查询／提交您的信用信息，并将视为已阅读并理解本协议的全部内容。\n" +
            "一、基本信息\n" +
            "为了保护您的个人信息，本协议在向您展示时将部分隐藏您的身份信息，您的身份信息详情以指定拿下钱包账户对应的实名信息为准。\n" +
            "用户姓名： \t\t\t身份证号码：1****************4\n" +
            "二、授权条款\n" +
            "您理解并同意，在您注册或使用拿下钱包时，本公司有权依据《征信业管理条例》及相关法律法规，向第三方支付/征信/金融机构或通过上述机构向其他拥有合法资质的第三方（包括但不限于具有资质的征信机构，运营商及其代理商、关联公司，公安部身份信息查询中心等）合法了解、获取、核实您的信用信息。本公司所获取的您的个人信用信息仅在拿下钱包业务中使用，且本公司保证不向其他机构、个人提供或披露。并同意本公司及上述机构将查询获取的信息进行保存、整理、加工，并用于评价本人信用情况或核实本人信息的真实性。但法律、法规、监管政策禁止查询的除外。\n" +
            "\n" +
            "您理解并同意，在您注册或使用拿下钱包时，本公司有权依据《征信业管理条例》及相关法律法规，向第三方支付/征信/金融机构提交您在本公司业务中产生的相关信息，包括但不限于个人基本信息、信用／提现信息、借款申请信息、借款合同信息以及还款行为信息、以及其他行为数据等，并记录在上述机构建设的个人信用信息数据库中。\n" +
            "您理解并同意，在您注册或使用拿下钱包时，本公司有权按照国家实名制要求，联合合作伙伴核实您的真实身份，您同意并授权身份验证服务方腾讯云计算（北京）有限责任公司及为此提供必要技术支持的深圳前海微众银行股份有限公司，在为您提供身份验证服务时可收集并存储您所提供的姓名，身份证信息及影像、银行卡信息及影像、设备信息等身份验证过程中产生的个人信息数据，以便识别您的真实身份，便遵守国家用户实名制相关的法律法规规定。\n" +
            "您理解并同意，在您通过拿下钱包平台与第三方金融机构（包括但不限于银行、消费金融公司、小额贷款公司等，简称“服务方”）达成交易意向时，您授权服务方可以自行或通过关联方通过中国人民银行个人信用信息基础数据库查询并使用您的信用报告，并同意其有权将信用活动中形成的交易记录等个人信贷交易信息及其他相关信息向中国人民银行个人信用信息基础数据库报送。\n" +
            "您理解并同意，在您通过拿下钱包平台与第三方金融机构（包括但不限于银行、消费金融公司、小额贷款公司等，简称“服务方”）达成交易意向时，您授权第三方金融机构可以向第三方支付/征信/金融机构或通过上述机构向其他拥有合法资质的第三方（包括但不限于具有资质的征信机构，运营商及其代理商、关联公司，公安部身份信息查询中心等）合法了解、获取、核实您的信用信息。第三方金融机构所获取的您的个人信用信息仅在拿下钱包业务中使用，且其保证不向其他机构、个人提供或披露。并同意上述机构将查询获取的信息进行保存、整理、加工，并用于评价本人信用情况或核实本人信息的真实性。但法律、法规、监管政策禁止查询的除外。\n" +
            "您理解并同意，如您使用了本公司合作的第三方支付机构提供的提现或支付功能，则应当依据《借款协议》或其他文件或协议中约定之日期按期进行还款，本公司有权通过短信、电话、社交账号等途径对您进行服务与还款提醒。您理解并同意，如您未有按期还款，您的个人逾期信息将可能向第三方（包括但不限于征信机构、金融机构、公众媒体等）进行分享或公布，记录在您的个人信用信息数据库中，对此造成的不良后果您已完全知悉。请您珍视自身的信用记录，按时履行还款义务。\n" +
            "您理解并同意，您应当对您的信息变动履行及时告知义务。若您的联系方式发生变化，应及时通知本公司，否则由此带来的影响和损失，本公司不承担责任。\n" +
            "本协议所提及的与本公司存在合作关系的第三方征信机构、金融机构包括但不限于：中诚信征信有限公司、考拉征信公司、淘宝、支付宝、芝麻信用管理有限公司、中智诚征信有限公司及未来开展合作的其他征信机构。\n" +
            "三、保密条款\n" +
            "（一）本平台重视对用户隐私的保护。因收集您的信息是出于遵守国家法律法规的规定以及向您提供服务及提升服务质量的目的，本公司对您的信用信息承担保密义务，不会为满足第三方的营销目的而向其出售或出租您的任何信息。\n" +
            "（二）本公司会在下列情况下才将您的个人信用信息与第三方共享：\n" +
            "1、事先获得用户的明确授权；\n" +
            "2、某些情况下，只有共享您的信息，才能提供您需要的服务和（或）产品，或处理您与他人的交易纠纷或争议；\n" +
            "3、某些服务和（或）产品由本公司的合作伙伴提供或由本公司与合作伙伴共同提供，本公司会与其共享提供服务和（或）产品需要的信息；\n" +
            "4、本公司与第三方进行联合推广活动，本公司可能与其共享活动过程中产生的、为完成活动所必要的个人信息，如参加活动的中奖名单、中奖人联系方式等，以便第三方能及时向您发放奖品；\n" +
            "5、为维护本公司和其他拿下钱包用户的合法权益；\n" +
            "6、根据法律规定及合理商业习惯，在本公司计划与其他公司合并或被其收购或进行其他资本市场活动（包括但不限于首次公开发行股票，债券发行）时，以及其他情形下本公司需要接受来自其他主体的尽职调查时，本公司会把您的信息提供给必要的主体，但本公司会通过和这些主体签署保密协议等方式要求其对您的个人信息采取合理的保密措施。根据有关的法律法规要求;\n" +
            "7、按照相关政府主管部门的要求；及\n" +
            "8、您理解并同意，我们可以储存您授权的原始信息；在您和我们的合作存续期间，我们随时可以重新采集和更新数据，对于经过加工和脱敏处理的数据，我们可以永久保存在服务器上。\n" +
            "（三）由于下列情况导致的个人信息泄露，本公司将不承担任何责任\n" +
            "1、由于用户将其用户密码告知他人或与他人共享注册账户与密码，由此导致的任何个人信息的泄漏，或其他非因本公司原因（因本协议第二条第二款所列情形导致信息泄漏的，不视为本公司原因）导致用户个人信息的泄露；\n" +
            "2、任何由于黑客攻击、电脑病毒入侵及其他不可抗力事件所导致的用户信息泄露、公开。\n" +
            "四、用户义务\n" +
            "（一）您保证，您所提供的个人信息均为您本人的真实信息，不可为他人的信息或虚假信息，若您涉嫌恶意信息作假或盗用他人信息，您的行为将可能记入网络征信系统，影响您的征信记录，同时本公司将保留追究您相应法律责任的权利。\n" +
            "（二）如您所提供的个人信息中的全部或部分信息为他人信息或虚假信息，本公司将有权暂停或终止与您的全部或部分服务，由此行为所产生的全部法律责任将由您承担，本公司将不对此承担法律责任。\n";

    public static void main2(String[] args) throws Exception {

        File file = new File("D:\\workspace\\IDEA_workspace\\MSJApiDemo\\src\\main\\resources\\aa.txt");

        //FileUtils.cover(file);
        //
        byte[] bytes = FileUtils.fileToBytes(file);
        //
        //     byte[] bytes1 = FileUtils.stringToBytes(x);
        //
        //     System.out.println("main(): ----------");

        Map<String, String> map = new HashMap<String, String>();
        //
        map.put("${name}", "seawater");
        map.put("${idNo}", "0000-0000");
        String srcPath = "D:\\workspace\\IDEA_workspace\\MSJApiDemo\\src\\main\\resources\\aa2.docx";
        String destPath = "D:\\workspace\\IDEA_workspace\\MSJApiDemo\\src\\main\\resources\\dd2.docx";
        docxSearchAndReplace(srcPath, destPath, map);
        //String destPath8 = "D:\\workspace\\IDEA_workspace\\MSJApiDemo\\src\\main\\resources\\dd6.doc";
        //searchAndReplace(destPath, destPath8, map);

        //The supplied data appears to be in the OLE2 Format. You are calling the part of POI that
        // deals with OOXML (Office Open XML) Documents. You need to call a different part of POI
        // to process this data (eg HSSF instead of XSSF)



        //if (url.endsWith(".docx")) {
        //    converter = new DocxToPDFConverter(inputStream, outputStream, true, true);
        //    converter.convert();
        //    fileInputStream = new FileInputStream(file);
        //} else if (url.endsWith(".doc")) {
        //String destPath1 = "D:\\workspace\\IDEA_workspace\\MSJApiDemo\\src\\main\\resources\\dd2.doc";
        String destPath2 = "D:\\workspace\\IDEA_workspace\\MSJApiDemo\\src\\main\\resources\\ddf.pdf";

        //The supplied data appears to be in the OLE2 Format. You are calling the part of POI that
        // deals with OOXML (Office Open XML) Documents. You need to call a different part of POI
        // to process this data (eg HSSF instead of XSSF)

        //The supplied data appears to be in the Office 2007+ XML. You are calling the part of POI
        // that deals with OLE2 Office Documents. You need to call a different part of POI
        // to process this data (eg XSSF instead of HSSF)

        InputStream inputStream = new FileInputStream(destPath);
        OutputStream outputStream = new FileOutputStream(destPath2);
        DocxToPDFConverter converter = new DocxToPDFConverter(inputStream, outputStream, true, true);
        converter.convert();
    }

    public static void docxSearchAndReplace(String srcPath, String destPath, Map<String, String> map) {
        // 操作 docx文件 不只是单纯文件名doc
        try {

            XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(srcPath));

            /**
             * 替换段落中的指定文字
             */
            Iterator<XWPFParagraph> itPara = document.getParagraphsIterator();
            while (itPara.hasNext()) {
                XWPFParagraph paragraph = (XWPFParagraph) itPara.next();
                Set<String> set = map.keySet();
                Iterator<String> iterator = set.iterator();
                while (iterator.hasNext()) {
                    String key = iterator.next();
                    List<XWPFRun> run = paragraph.getRuns();
                    for (int i = 0; i < run.size(); i++) {
                        if (run.get(i).getText(run.get(i).getTextPosition()) != null &&
                                run.get(i).getText(run.get(i).getTextPosition()).equals(key)) {
                            /**
                             * 参数0表示生成的文字是要从哪一个地方开始放置,设置文字从位置0开始
                             * 就可以把原来的文字全部替换掉了
                             *
                             * 需要注意的是在模板文件中,我们定义的标识左右都需要加空格，否则可能会出现无法替换的情况。
                             */
                            run.get(i).setText(map.get(key), 0);
                        }
                    }
                }
            }

            /**
             * 替换表格中的指定文字
             */
            Iterator<XWPFTable> itTable = document.getTablesIterator();
            while (itTable.hasNext()) {
                XWPFTable table = (XWPFTable) itTable.next();
                int count = table.getNumberOfRows();
                for (int i = 0; i < count; i++) {
                    XWPFTableRow row = table.getRow(i);
                    List<XWPFTableCell> cells = row.getTableCells();
                    for (XWPFTableCell cell : cells) {
                        for (Map.Entry<String, String> e : map.entrySet()) {
                            if (cell.getText().equals(e.getKey())) {
                                cell.removeParagraph(0);
                                cell.setText(e.getValue());
                            }
                        }
                    }
                }
            }
            FileOutputStream outStream = null;
            outStream = new FileOutputStream(destPath);
            document.write(outStream);
            outStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }



}
