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
    private int allPdfNum = 0;
    private int allExcelRowNum = 0;
    private int allErrorNum = 0;
    private StringBuffer allErrorFileName = new StringBuffer();
    private static final int THREAD_NUM = 4;

    private synchronized void setAllPdfNum(int pdfNum){
        this.allPdfNum += pdfNum;
    }

    private synchronized void setAllErrorNum(int errorNum){
        this.allErrorNum += errorNum;
    }

    private synchronized void setAllErrorFileName(String errorFileName){
        this.allErrorFileName.append(errorFileName);
    }

    // 根据选择的是文件或目录选择不同处理方式
    public void fileOrDirectory(String inputPath, String outputPath) {
        File inputFile = new File(inputPath);
        File outputFile = new File(outputPath);
        // 如果是文件夹，则创建Excel文件
        if (outputFile.isDirectory()) {
            SimpleDateFormat ft = new SimpleDateFormat("yyyyMMddhhmmss");
            outputPath += "\\电子发票" + ft.format(new Date()) + ".xlsx";
            try {
                createExcel(outputPath);
            } catch (IOException e) {
                JOptionPane.showMessageDialog(null, "创建Excel文件失败！请稍后再试。", "错误", 0);
            }
        }
        try {
            // 如果是文件，直接读写
            if (inputFile.isFile()) {
                List<HashMap<String, String>> mapList = new ArrayList<>();
                try {
                    mapList.add(new PDFUtil().readPdf(inputPath));
                    allPdfNum += 1;
                } catch (Exception e) {
                    JOptionPane.showMessageDialog(null, "PDF文件打开错误！请检查后重试。", "错误", 0);
                    return;
                }
                writeExcel(mapList, outputPath);
            // 如果是文件夹，对目录下所有文件多线程读写
            } else if (inputFile.isDirectory()) {
                String[] fileNameList = inputFile.list();
                readByMultThread(fileNameList, inputPath, outputPath);
            }
        } catch (IOException e) {
            JOptionPane.showMessageDialog(null, "Excel文件打开错误！请关闭Excel文件后重试。", "错误", 0);
            return;
        }
        JOptionPane.showMessageDialog(null, "转换完成！" + "\n读取PDF文件数：" + allPdfNum + "\n写入Excel行数：" + allExcelRowNum
                + "\n错误数：" + allErrorNum + "\n错误文件：\n" + (allErrorFileName.length() == 0 ? "无" : allErrorFileName), "完成", JOptionPane.PLAIN_MESSAGE);
    }

    // 多线程读pdf
    private void readByMultThread(String[] fileNameList, String inputPath, String outputPath) throws IOException {
        // 分配文件数
        int increment = fileNameList.length / THREAD_NUM + 1;
        int start = 0;
        List<Thread> readThreadList = new ArrayList();
        for (int i = 0; i < THREAD_NUM; i++) {
            if ((start + increment) > fileNameList.length) {
                increment = fileNameList.length - start;
            }
            readThreadList.add(new Thread(new ReadThread(fileNameList, inputPath, outputPath, start, start + increment)));
            start += increment;
        }
        // 启动线程并挂靠
        for (Thread thread : readThreadList) {
            thread.start();
        }
        for (Thread thread : readThreadList) {
            try {
                thread.join();
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }
    }

    // 线程类，重写run()
    class ReadThread implements Runnable{
        private String[] fileNameList;
        private String pdfPath;
        private String excelPath;
        private int startnum;
        private int endnum;
        private int pdfNum = 0;
        private int errorNum = 0;
        private StringBuffer errorFileName = new StringBuffer("");

        public ReadThread(String[] fileNameList,String pdfPath, String excelPath, int startnum, int endnum){
            this.fileNameList = fileNameList;
            this.pdfPath = pdfPath;
            this.excelPath = excelPath;
            this.startnum = startnum;
            this.endnum = endnum;
        }

        @Override
        public void run() {
            List<HashMap<String, String>> mapList = new ArrayList<>();

            for (int i = startnum; i < endnum; i++) {
                //System.out.println(i); //+ Thread.currentThread().getName());
                String fileName = fileNameList[i];
                if("pdf".equals(fileName.substring(fileName.lastIndexOf(".") + 1))) {
                    pdfNum += 1;
                    try {
                        HashMap<String, String> map = new PDFUtil().readPdf(pdfPath + "\\" + fileName);
                        mapList.add(map);
                    }catch (Exception e){
                        errorNum += 1;
                        errorFileName.append(fileName + "\n");
                    }
                }
            }
            setAllPdfNum(pdfNum);
            setAllErrorNum(errorNum);
            setAllErrorFileName(errorFileName.toString());
            try {
                writeExcel(mapList, excelPath);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


    // 写入excel
    private synchronized void writeExcel (List < HashMap < String, String >> mapList, String excelPath) throws
    IOException {
        Workbook wb = null;
        FileInputStream fileInput = new FileInputStream(excelPath);
        if ("xls".equals(excelPath.substring(excelPath.lastIndexOf(".") + 1))) {
            wb = new HSSFWorkbook(fileInput);
        } else {
            wb = new XSSFWorkbook(fileInput);
        }
        Sheet sheet = wb.getSheetAt(0);
        Row head = sheet.getRow(0);
        for (HashMap<String, String> map : mapList) {
            Row newRow = sheet.createRow(sheet.getLastRowNum() + 1);
            for (int i = 0; i < head.getPhysicalNumberOfCells(); i++) {
                Cell newCell = newRow.createCell(i);
                String key = head.getCell(i).toString();
                newCell.setCellValue(map.get(key));
            }
            allExcelRowNum += 1;
        }
        FileOutputStream fileOutputStream = new FileOutputStream(excelPath);
        wb.write(fileOutputStream);
        fileOutputStream.close();
    }

    // 创建excel
    private void createExcel (String outputPath) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("电子发票");
        XSSFRow row = sheet.createRow(0);
        // 首行
        String[] key = {"发票类型", "机器编号", "发票代码", "发票号码", "开票日期", "校验码",
                "购买方名称", "购买方纳税人识别号", "购买方地址、电话", "购买方开户行及账号",
                "密码区",
                "项目名称", "车牌号", "类型", "通行日期起", "通行日期止", "金额", "税率", "税额",
                "价税合计（大写）", "税价合计",
                "开票抬头", "销售方纳税人识别号", "销售方地址、电话", "销售方开户行及账号",
                "备注",
                "收款人", "复核", "开票人"};
        for (int i = 0; i < key.length; i++) {
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

        FileOutputStream output = new FileOutputStream(outputPath);
        wb.write(output);
        output.flush();
    }
}
