import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

/**
 * @author Duckbill-lujixuan
 * @date 2019/01/08
 */
public class GuiUtil {
    private JLabel inputLabel = new JLabel("读取：");
    private JTextField inputField = new JTextField(25);
    private JButton inputButton = new JButton("浏览");
    private JLabel outpubLabel = new JLabel("输出：");
    private JTextField outputField = new JTextField(25);
    private JButton outputButton = new JButton("浏览");
    private JButton convensionButton = new JButton("转换");
    //得到系统桌面文件夹 Desktop
    FileSystemView homeFileSystemView = FileSystemView.getFileSystemView();
    File homePath = homeFileSystemView.getHomeDirectory();

    public GuiUtil(){
        JFrame jFrame = new JFrame("电子发票转Excel");
        JPanel panel1 = new JPanel();
        JPanel panel2 = new JPanel();
        JPanel panel3 = new JPanel();
        jFrame.setLayout(new GridLayout(3,3));
        // 设置居于屏幕中央
        jFrame.setLocationRelativeTo(null);
        outputField.setText(homePath.getPath());
        inputField.setDisabledTextColor(Color.black);
        outputField.setDisabledTextColor(Color.black);
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

    //选择读取文件按钮
    class InputActionListener implements ActionListener{
        @Override
        public void actionPerformed(ActionEvent arg){
            JFileChooser fc = new JFileChooser();
            fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
            fc.setCurrentDirectory(homePath);
            fc.setFileFilter(new FileNameExtensionFilter("pdf(*.pdf)", "pdf"));
            int val = fc.showOpenDialog(null);
            if (val == JFileChooser.APPROVE_OPTION) {
                String path = fc.getSelectedFile().toString();
                if(pdfOrDirectory(path)){
                    inputField.setText(path);
                }
            }
        }
    }

    //选择输入文件按钮
    class OutputActionListener implements ActionListener{
        @Override
        public void actionPerformed(ActionEvent arg){
            JFileChooser fc = new JFileChooser();
            fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
            fc.setCurrentDirectory(new File("C:\\Users\\Administrator\\Desktop"));
            fc.setFileFilter(new FileNameExtensionFilter("xls(*.xls, *.xlsx)", "xls", "xlsx"));
            int val = fc.showOpenDialog(null);
            if(val == JFileChooser.APPROVE_OPTION) {
                String path = fc.getSelectedFile().toString();
                if(excelOrDirectory(path)){
                    outputField.setText(path);
                }
            }
        }
    }

    //转换按钮
    class ConvensionActionListener implements  ActionListener{
        @Override
        public void actionPerformed(ActionEvent arg){
            String inputPath = inputField.getText();
            String outputPath = outputField.getText();
            if(pdfOrDirectory(inputPath) && excelOrDirectory(outputPath)){
                new ExcelUtil().fileOrDirectory(inputPath, outputPath);
            }
        }
    }

    public boolean pdfOrDirectory(String path){
        File file = new File(path);
        if(file.isFile()){
            if("pdf".equals(path.substring(path.lastIndexOf(".") + 1))){
                return true;
            }
        }else if(file.isDirectory()){
            return true;
        }
        JOptionPane.showMessageDialog(null, "请选择一个PDF文件或文件夹！", "错误",0);
        return false;
    }

    public boolean excelOrDirectory(String path){
        File file = new File(path);
        if(file.isFile()){
            if("xls".equals(path.substring(path.lastIndexOf(".") + 1)) || "xlsx".equals(path.substring(path.lastIndexOf(".") + 1))){
                return true;
            }
        }else if(file.isDirectory()){
            return true;
        }
        JOptionPane.showMessageDialog(null, "请选择一个Excel文件或文件夹！", "错误",0);
        return false;
    }

    public static void main(String[] args){
        new GuiUtil();
    }
}
