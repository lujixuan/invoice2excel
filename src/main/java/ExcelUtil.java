import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

/**
 * @author Duckbill-lujixuan
 * @date 2019/01/08
 */
public class ExcelUtil {
    public int PDF_NUM = 0;
    public int EXCEL_ROW_NUM = 0;
    public int ERROR_NUM = 0;
    public StringBuffer ERROR_FILE_NAME = new StringBuffer();

    public void writeExcel(List<HashMap<String, String>> mapList, String excelPath) throws IOException {
        Workbook wb = null;
        FileInputStream fileInput = new FileInputStream(excelPath);
        if("xls".equals(excelPath.substring(excelPath.lastIndexOf(".") + 1))){
            wb = new HSSFWorkbook(fileInput);
        }else{
            wb = new XSSFWorkbook(fileInput);
        }
        Sheet sheet = wb.getSheetAt(0);
        Row head = sheet.getRow(0);
        for(HashMap<String, String> map:mapList) {
            Row newRow = sheet.createRow(sheet.getLastRowNum() + 1);
            for (int i = 0; i < head.getPhysicalNumberOfCells(); i++) {
                Cell newCell = newRow.createCell(i);
                String key = head.getCell(i).toString();
                newCell.setCellValue(map.get(key));
            }
            EXCEL_ROW_NUM += 1;
        }
        FileOutputStream fileOutputStream = new FileOutputStream(excelPath);
        wb.write(fileOutputStream);
        fileOutputStream.close();
    }

    public void fileOrDirectory(String inputPath, String outputPath){
        File inputFile = new File(inputPath);
        File outputFile = new File(outputPath);
        // 如果是文件夹，则创建Excel文件
        if(!outputFile.isFile()){
            SimpleDateFormat ft = new SimpleDateFormat("yyyyMMddhhmmss");
            outputPath += "\\电子发票" + ft.format(new Date()) +".xlsx";
            try {
                createExcel(outputPath);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        try {
            if (inputFile.isFile()) {
                List<HashMap<String, String>> mapList = new ArrayList<>();
                try {
                    mapList.add(new PDFUtil().readPdf(inputPath));
                    PDF_NUM += 1;
                }catch (Exception e){
                    JOptionPane.showMessageDialog(null, "PDF文件打开错误！请检查后重试。", "错误",0);
                    return;
                }
                writeExcel(mapList, outputPath);
            } else if (inputFile.isDirectory()) {
                String[] fileNameList = inputFile.list();
                List<HashMap<String, String>> mapList = new ArrayList<>();
                // 识别所有pdf文件，添加入Hashmap中
                for (String fileName : fileNameList) {
                    if("pdf".equals(fileName.substring(fileName.lastIndexOf(".") + 1))) {
                        PDF_NUM += 1;
                        try {
                            HashMap<String, String> map = new PDFUtil().readPdf(inputPath + "\\" + fileName);
                            mapList.add(map);
                        }catch (Exception e){
                            ERROR_NUM += 1;
                            ERROR_FILE_NAME.append(fileName + "\n");
                        }
                    }
                }
                writeExcel(mapList, outputPath);
            }
        }catch (IOException e){
            JOptionPane.showMessageDialog(null, "Excel文件打开错误！请关闭Excel文件后重试。", "错误",0);
            return;
        }
        JOptionPane.showMessageDialog(null, "转换完成！" + "\n读取PDF文件数：" + PDF_NUM + "\n写入Excel行数：" + EXCEL_ROW_NUM
                + "\n错误数：" + ERROR_NUM + "\n错误文件：\n" + (ERROR_FILE_NAME.length() == 0 ? "无" : ERROR_FILE_NAME), "完成",JOptionPane.PLAIN_MESSAGE);
    }

    public void createExcel(String excelPath) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("电子发票");
        XSSFRow row = sheet.createRow(0);
        String[] key = {"发票类型", "机器编号", "发票代码", "发票号码", "开票日期", "校验码",
                "购买方名称", "购买方纳税人识别号", "购买方地址、电话","购买方开户行及账号",
                "密码区",
                "项目名称", "车牌号", "类型", "通行日期起", "通行日期止", "金额", "税率", "税额",
                "价税合计（大写）","税价合计",
                "开票抬头", "销售方纳税人识别号", "销售方地址、电话", "销售方开户行及账号",
                "备注",
                "收款人", "复核", "开票人"};
        for(int i = 0; i < key.length; i++){
            Cell cell = row.createCell(i);
            cell.setCellValue(key[i]);
        }
        // 设置列宽
        sheet.setColumnWidth(0, 22 * 256);
        sheet.setColumnWidth(1, 14 * 256);
        sheet.setColumnWidth(2, 14 * 256);
        sheet.setColumnWidth(4, 16 * 256);
        sheet.setColumnWidth(6, 28 * 256);
        sheet.setColumnWidth(7, 20 * 256);
        sheet.setColumnWidth(11, 21 * 256);
        sheet.setColumnWidth(14, 11 * 256);
        sheet.setColumnWidth(15, 11 * 256);
        sheet.setColumnWidth(19, 16 * 256);
        sheet.setColumnWidth(21, 22 * 256);
        sheet.setColumnWidth(22, 20 * 256);
        sheet.setColumnWidth(23, 42 * 256);
        sheet.setColumnWidth(24, 42 * 256);
        // 首行不滚动
        sheet.createFreezePane(0, 1);

        FileOutputStream output = new FileOutputStream(excelPath);
        wb.write(output);
        output.flush();
    }
}
