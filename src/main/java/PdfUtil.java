import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import java.io.File;

public class PdfUtil {

    public void readPdf(String filePath){
        PDDocument doc = null;
        String content = "";
        try {
            //加载一个pdf对象
            doc =PDDocument.load(new File(filePath));
            //获取一个PDFTextStripper文本剥离对象
            PDFTextStripper textStripper =new PDFTextStripper();
            content=textStripper.getText(doc);
            System.out.println("内容:"+content);
            //关闭文档
            doc.close();
        } catch (Exception e) {
            System.out.println("pdf转化出错！");
        }
    }

    public static void main(String[] args){
        String pdfPath = "C:\\Users\\Administrator\\Documents\\WXWork\\1688853865604551\\Cache\\File\\2019-01\\鲁通卡发票\\91310110734550773T_f7bb55c1644141ab9e9b318b90cc5efe.pdf";
        PdfUtil pdfUtil = new PdfUtil();
        pdfUtil.readPdf(pdfPath);
    }

}
