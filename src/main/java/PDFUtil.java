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
        //设置识别窗口大小
        Rectangle header = new Rectangle(10, 7, 762, 80);
        Rectangle purchaser = new Rectangle(103, 90, 220, 60);
        Rectangle password = new Rectangle(350,80,306,70);
        Rectangle tax = new Rectangle(10, 160, 762, 40);
        Rectangle sum = new Rectangle(120, 250, 700, 50);
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
            // 将窗口添加入stripper
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
            // 根据识别窗口读取数据
            passwordString = stripper.getTextForRegion("password");
            headerString = stripper.getTextForRegion("header");
            purchaserString = stripper.getTextForRegion("purchaser");
            taxString = stripper.getTextForRegion("tax");
            sumString = stripper.getTextForRegion("sum");
            sellerString = stripper.getTextForRegion("seller");
            remarkString = stripper.getTextForRegion("remark");
            endString = stripper.getTextForRegion("end");
            // 读取完成，关闭文档
            doc.close();
        } catch (Exception e) {
            System.out.println("pdf转化出错！");
        }
        // 分割字符串，将数据存入HashMap
        passwordString = passwordString.replace("\r\n", "").replace(" ", "");
        remarkString = remarkString.replace("\n","").replace("\r","");
        String[] endArray = endString.replace("\r\n", " ").split("[: "+" ]");
        String[] sellerArray = sellerString.replace("\r\n", " ").split("[: "+" ]");
        String[] taxArray = taxString.replace("\r\n", " ").split("[: "+" ]");
        String[] purchaserArray = purchaserString.replace("\r\n", " ").split("[: "+" ]");
        String[] headerArray = headerString.replace("\r\n", " ").split("[: "+" ]");
        String[] sumArray = sumString.replace("\r\n"," ").replace("￥","").split("[:" + "￥" + " ]");
        map.put("密码区", passwordString);
        map.put("价税合计（大写）", sumArray[3]);
        map.put("税价合计", sumArray[4]);
        map.put("备注", remarkString);
        map.put("收款人", endArray[10]);
        map.put("复核", endArray[11]);
        map.put("开票人", endArray[12]);
        map.put("开票抬头", sellerArray[0]);
        map.put("销售方纳税人识别号", sellerArray[1]);
        map.put("销售方地址、电话", sellerArray[2]);
        map.put("销售方开户行及账号", sellerArray[3]);
        map.put("项目名称", taxArray[0]);
        map.put("车牌号", taxArray[1]);
        map.put("类型", taxArray[2]);
        map.put("通行日期起", taxArray[3]);
        map.put("通行日期止", taxArray[4]);
        map.put("金额", sumArray[1]);
        map.put("税率", taxArray[6]);
        map.put("税额", sumArray[2]);
        map.put("购买方名称", purchaserArray[0]);
        map.put("购买方纳税人识别号", purchaserArray[1]);
        map.put("购买方地址、电话", purchaserArray.length >= 3 ? purchaserArray[2] : "");
        map.put("购买方开户行及账号", purchaserArray.length >= 4 ? purchaserArray[3] : "");
        map.put("发票类型", headerArray[0]);
        map.put("机器编号", headerArray[12]);
        map.put("发票代码", headerArray[13]);
        map.put("发票号码", headerArray[14]);
        map.put("开票日期", headerArray[15]);
        map.put("校验码", headerArray[16]+headerArray[17]+headerArray[18]+headerArray[19]);

        return map;
    }

    public void stringIsNull(String stringName){
        // TODO: 字符串数组不存在该下标时，输出错误
    }

    public void stringIsEmpty(String stringName){
        // TODO: 字符串数组该下标字符为空时，输出错误
    }
}
