import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class GuiUtil {
    private JLabel inputLabel = new JLabel("输入：");
    private JTextField inputField = new JTextField(25);
    private JButton inputButton = new JButton("浏览");
    private JLabel outpubLabel = new JLabel("输出：");
    private JTextField outputField = new JTextField(25);
    private JButton outputButton = new JButton("浏览");
    private JButton convensionButton = new JButton("转换");
    private JTextArea convensionTextArea = new JTextArea(5,40);

    public GuiUtil(){
        JFrame jFrame = new JFrame("电子发票转Excel");
        JPanel panel1 = new JPanel();
        JPanel panel2 = new JPanel();
        JPanel panel3 = new JPanel();
        jFrame.setLayout(new GridLayout(3,3));
        // 设置居于屏幕中央
        jFrame.setLocationRelativeTo(null);
        inputField.setText("选择一个文件夹或PDF文件");
        outputField.setText("选择一个文件夹或Excel文件");
        inputField.setEnabled(false);
        outputField.setEnabled(false);
        panel1.add(inputLabel);
        panel1.add(inputField);
        panel1.add(inputButton);
        panel2.add(outpubLabel);
        panel2.add(outputField);
        panel2.add(outputButton);
        panel3.add(convensionButton);
        jFrame.add(panel1);
        jFrame.add(panel2);
        jFrame.add(panel3);
        jFrame.pack();
        jFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        inputButton.addActionListener(new InputActionListener());
        outputButton.addActionListener(new OutputActionListener());
        convensionButton.addActionListener(new ConvensionActionListener());
        jFrame.setVisible(true);
    }

    class InputActionListener implements ActionListener{
        @Override
        public void actionPerformed(ActionEvent arg){
            JFileChooser fc = new JFileChooser();
            fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
            fc.setCurrentDirectory(new File("C:\\Users\\Administrator\\Desktop"));
            fc.setFileFilter(new FileNameExtensionFilter("pdf(*.pdf)", "pdf"));
            int val = fc.showOpenDialog(null);
            if(val == JFileChooser.APPROVE_OPTION)
            {
                String path = fc.getSelectedFile().toString();
                File file = new File(path);
                if(file.isFile()){
                    if(!"pdf".equals(path.substring(path.lastIndexOf(".") + 1))){
                        JOptionPane.showMessageDialog(null, "请选择PDF文件或文件夹！", "错误",0);
                        return;
                    }
                }
                inputField.setText(fc.getSelectedFile().toString());
            }
        }
    }

    class OutputActionListener implements ActionListener{
        @Override
        public void actionPerformed(ActionEvent arg){
            JFileChooser fc = new JFileChooser();
            fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
            fc.setCurrentDirectory(new File("C:\\Users\\Administrator\\Desktop"));
            fc.setFileFilter(new FileNameExtensionFilter("xls(*.xls, *.xlsx)", "xls", "xlsx"));
            int val = fc.showOpenDialog(null);
            if(val == JFileChooser.APPROVE_OPTION)
            {
                String path = fc.getSelectedFile().toString();
                File file = new File(path);
                if(file.isFile()){
                    if(!"xls".equals(path.substring(path.lastIndexOf(".") + 1)) && !"xlsx".equals(path.substring(path.lastIndexOf(".") + 1))){
                        JOptionPane.showMessageDialog(null, "请选择Excel文件或文件夹！", "错误",0);
                        return;
                    }
                }
                outputField.setText(fc.getSelectedFile().toString());
            }
        }
    }

    class ConvensionActionListener implements  ActionListener{
        @Override
        public void actionPerformed(ActionEvent arg){
            String inputPath = inputField.getText();
            String outputPath = outputField.getText();
            fileOrDirectory(inputPath, outputPath);
        }

        public void fileOrDirectory(String inputPath, String outputPath){
            File inputFile = new File(inputPath);
            File outputFile = new File(outputPath);
            if(!outputFile.isFile()){
                SimpleDateFormat ft = new SimpleDateFormat("yyyyMMddhhmmss");
                outputPath += "\\电子发票" + ft.format(new Date()) +".xlsx";
                try {
                    new ExcelUtil().createExcel(outputPath);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

            String excelPath = outputPath;
            FileInputStream fileInput = null;

            try {
                fileInput = new FileInputStream(outputPath);
            } catch (IOException e) {
                JOptionPane.showMessageDialog(null, "文件打开错误！请关闭Excel文件后重试。", "错误",0);
                return;
            }
            if(inputFile.isFile()){
                writeIntoFile(inputPath, fileInput, excelPath);
            }else if (inputFile.isDirectory()){
                String[] fileNameList = inputFile.list();
                for(String fileName:fileNameList){
                    writeIntoFile(inputPath + "\\" + fileName, fileInput, excelPath);
                }
            }
            JOptionPane.showMessageDialog(null, "转换完成！", "成功",JOptionPane.PLAIN_MESSAGE);
        }

        public void writeIntoFile(String inputPath, FileInputStream inputStream, String excelPath){
            String inputType = inputPath.substring(inputPath.lastIndexOf(".") + 1);
            if("pdf".equals(inputType)){
                try {
                    new ExcelUtil().writeExcel(inputPath, inputStream, excelPath);
                } catch (IOException e) {

                }
            }
        }
    }



    public static void main(String[] args){
        new GuiUtil();
    }
}
