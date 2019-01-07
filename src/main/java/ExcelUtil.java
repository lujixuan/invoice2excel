import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;

public class ExcelUtil {
    public void writeExcel(String excelPath, HashMap<String, String> map) throws IOException {
        File file = new File(excelPath);
        if(!file.isFile()){
            SimpleDateFormat ft = new SimpleDateFormat("yyyy-MM-dd-hh-mm-ss");
            excelPath += "\\电子发票" + ft.format(new Date()) +".xlsx";
            createExcel(excelPath);
        }
        Workbook wb = null;
        FileInputStream fileInput = null;
        try {
            fileInput = new FileInputStream(excelPath);
        } catch (IOException e) {
            e.printStackTrace();
        }
        if("xls".equals(excelPath.substring(excelPath.lastIndexOf(".") + 1))){
            wb = new HSSFWorkbook(fileInput);
        }else{
            wb = new XSSFWorkbook(fileInput);
        }
        Sheet sheet = wb.getSheetAt(0);
        Row head = sheet.getRow(0);
        Row newRow = sheet.createRow(sheet.getLastRowNum() + 1);
        for(int i = 0; i < head.getPhysicalNumberOfCells(); i++) {
            Cell newCell = newRow.createCell(i);
            String key = head.getCell(i).toString();
            newCell.setCellValue(map.get(key));
        }
        FileOutputStream fileOutputStream = new FileOutputStream(excelPath);
        wb.write(fileOutputStream);
        fileOutputStream.close();
    }

    public void createExcel(String excelPath) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("电子发票");
        XSSFRow row = sheet.createRow(0);
        String[] key = {"发票类型", "机器编号", "发票代码", "发票号码", "开票日期", "校验码", "购买方名称", "购买方纳税人识别号", "购买方地址、电话",
                "购买方开户行及账号", "密码区", "项目名称", "车牌号", "类型", "通行日期起", "通行日期止", "金额", "税率", "税额",
                "价税合计（大写）", "销售方名称", "销售方纳税人识别号", "销售方地址、电话", "销售方开户行及账号", "备注", "收款人", "复核", "开票人"};
        for(int i = 0; i < key.length; i++){
            Cell cell = row.createCell(i);
            cell.setCellValue(key[i]);
        }
        sheet.setColumnWidth(0,22*256);
        sheet.setColumnWidth(2,13*256);
        sheet.setColumnWidth(4,15*256);
        sheet.setColumnWidth(6,26*256);
        sheet.setColumnWidth(20,27*256);
        FileOutputStream output = new FileOutputStream(excelPath);
        wb.write(output);
        output.flush();
    }

    public static void main(String[] args) throws IOException {
        String pdfPath = "C:\\Users\\Administrator\\Desktop\\发票\\111.pdf";
        String excelPath = "C:\\Users\\Administrator\\Desktop\\发票";

        PDFUtil pdfFile = new PDFUtil();
        ExcelUtil excelFile = new ExcelUtil();
        excelFile.writeExcel(excelPath, pdfFile.readPdf(pdfPath));
    }
}
