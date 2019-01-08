import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.text.PDFTextStripperByArea;

import java.awt.*;
import java.io.File;
import java.util.HashMap;

/**
 * @author Duckbill-lujixuan
 * @date 2019/01/08
 */
public class PDFUtil {

    public HashMap<String, String> readPdf(String readPath){
        Rectangle header = new Rectangle(10, 7, 762, 80);
        Rectangle purchaser = new Rectangle(103, 90, 220, 60);
        Rectangle password = new Rectangle(350,80,306,70);
        Rectangle tax = new Rectangle(10, 160, 762, 40);
        Rectangle sum = new Rectangle(120, 280, 200, 60);
        Rectangle seller = new Rectangle(103,300,220,60);
        Rectangle remark = new Rectangle(360,300,300,60);
        Rectangle end = new Rectangle(10,360,400,38);
        String passwordString = "";
        String headerString = "";
        String purchaserString = "";
        String taxString = "";
        String sumString = "";
        String sellerString = "";
        String remarkString = "";
        String endString = "";
        HashMap<String, String> map = new HashMap<>();
        try {
            PDDocument doc = PDDocument.load(new File(readPath));
            PDFTextStripperByArea stripper = new PDFTextStripperByArea();
            stripper.addRegion("header", header);
            stripper.addRegion("purchaser", purchaser);
            stripper.addRegion("password", password);
            stripper.addRegion("tax", tax);
            stripper.addRegion("sum", sum);
            stripper.addRegion("seller", seller);
            stripper.addRegion("remark", remark);
            stripper.addRegion("end", end);
            PDPageTree allPages = doc.getDocumentCatalog().getPages();
            PDPage firstPage = (PDPage)allPages.get(0);
            stripper.extractRegions( firstPage );

            passwordString = stripper.getTextForRegion("password");
            headerString = stripper.getTextForRegion("header");
            purchaserString = stripper.getTextForRegion("purchaser");
            taxString = stripper.getTextForRegion("tax");
            sumString = stripper.getTextForRegion("sum");
            sellerString = stripper.getTextForRegion("seller");
            remarkString = stripper.getTextForRegion("remark");
            endString = stripper.getTextForRegion("end");

            doc.close();
        } catch (Exception e) {
            System.out.println("pdf转化出错！");
        }
        passwordString = passwordString.replace('\r', ' ');
        passwordString = passwordString.replace('\n', ' ');
        passwordString = passwordString.replace(" ", "");
        String[] endArray = endString.split("[: \n\" + \" \r\" + \"  ]");
        String[] sellerArray = sellerString.split("[: \n\" + \" \r\" + \"  ]");
        String[] taxArray = taxString.split("[: \n\" + \" \r\" + \"  ]");
        String[] purchaserArray = purchaserString.split("[: \n" + " \r" + "  ]");
        String[] headerArray = headerString.split("[: \n \r  ]");
        map.put("密码区", passwordString.split("[: \n" + " \r" + "  ]")[0]);
        map.put("价税合计（大写）", sumString.split("[: \n" + " \r" + "  ]")[0]);
        map.put("备注", remarkString.split("[\r]")[0]);
        map.put("收款人", endArray[10]);
        map.put("复核", endArray[11]);
        map.put("开票人", endArray[12]);
        map.put("销售方名称", sellerArray[0]);
        map.put("销售方纳税人识别号", sellerArray[2]);
        map.put("销售方地址、电话", sellerArray[4]);
        map.put("销售方开户行及账号", sellerArray[6]);
        map.put("项目名称", taxArray[0]);
        map.put("车牌号", taxArray[1]);
        map.put("类型", taxArray[2]);
        map.put("通行日期起", taxArray[3]);
        map.put("通行日期止", taxArray[4]);
        map.put("金额", taxArray[5]);
        map.put("税率", taxArray[6]);
        map.put("税额", taxArray[7]);
        map.put("购买方名称", purchaserArray[0]);
        map.put("购买方纳税人识别号", purchaserArray[2]);
        map.put("购买方地址、电话", purchaserArray.length == 4 ? purchaserArray[4] : "");
        map.put("购买方开户行及账号", purchaserArray.length == 6 ? purchaserArray[6] : "");
        map.put("发票类型", headerArray[0]);
        map.put("机器编号", headerArray[17]);
        map.put("发票代码", headerArray[19]);
        map.put("发票号码", headerArray[21]);
        map.put("开票日期", headerArray[23]);
        map.put("校验码", headerArray[25]+headerArray[26]+headerArray[27]+headerArray[28]);

        return map;
    }
}
