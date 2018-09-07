package com.ycl.main;


import org.jb2011.lnf.beautyeye.BeautyEyeLNFHelper;

import javax.swing.*;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileSystemView;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.Objects;


/**
 * @author Yin Changlei
 * @dateTime 2018/9/7 15:35
 */


/**
 * @author Yin Changlei
 * @dateTime 2018/8/10 15:03
 */
public class App extends JFrame implements ActionListener {

    private JPanel jPanel;
    private JButton btnChooseFile;
    private JTextField txtCsvFile;
    private JButton btnStart;
    private JLabel lblInfo;

    public App() {
        this.setTitle("Csv Transfer To Excel");
        this.setSize(520, 260);
        this.setLocation(400, 200);
        this.setResizable(false);

        jPanel = new JPanel();
        jPanel.setLayout(null);

        btnChooseFile = new JButton("选择CSV文件");
        btnChooseFile.setBounds(10, 10, 100, 30);
        btnChooseFile.addActionListener(this);
        jPanel.add(btnChooseFile);

        txtCsvFile = new JTextField();
        txtCsvFile.setBounds(10, 50, 400, 30);
        txtCsvFile.setEnabled(false);
        jPanel.add(txtCsvFile);

        btnStart = new JButton("开始转换");
        btnStart.setBounds(10, 100, 100, 30);
        btnStart.setFont(Utils.myFont(14));
        btnStart.setForeground(Color.BLUE);
        btnStart.addActionListener(this);
        jPanel.add(btnStart);

        lblInfo = new JLabel("");
        lblInfo.setBounds(10, 140, 400, 30);
        lblInfo.setForeground(Color.RED);
        lblInfo.setFont(Utils.myFont(14));
        jPanel.add(lblInfo);

        this.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        this.setContentPane(jPanel);
    }


    /**
     * 获取文件后缀
     */
    private String getSuffix(String fileName) {
        return fileName.substring(fileName.lastIndexOf(".") + 1);
    }


    public static void main(String[] args) {
        try {
            BeautyEyeLNFHelper.launchBeautyEyeLNF();
            UIManager.put("RootPane.setupButtonVisible", false);
        } catch (Exception e) {
            e.printStackTrace();
        }
        App app = new App();
        app.setVisible(true);
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        lblInfo.setText("");
        if (e.getSource() == btnChooseFile) {
            String filePath = this.selectFile();
            txtCsvFile.setText(filePath);
        } else if (e.getSource() == btnStart) {
            String path = txtCsvFile.getText();
            if (path == null || path.trim().equals("")) {
                lblInfo.setText("请行选择Csv文件");
                return;
            }
            transferToExcel(path);
        }
    }

    /**
     * 选择文件
     */
    private String selectFile() {
        String path = "";
        JFileChooser fileChooser = new JFileChooser();
        FileSystemView fsv = FileSystemView.getFileSystemView();  //注意了，这里重要的一句
        //System.out.println(fsv.getHomeDirectory());                //得到桌面路径
        fileChooser.setCurrentDirectory(fsv.getHomeDirectory());
        fileChooser.setDialogTitle("请选择要转换的CSV文件");
        fileChooser.setApproveButtonText("确定");
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        FileFilter filter = new ExtensionFileFilter("CSV格式文件", new String[]{"csv"});
        fileChooser.setFileFilter(filter);
        int result = fileChooser.showOpenDialog(this);
        if (JFileChooser.APPROVE_OPTION == result) {
            path = fileChooser.getSelectedFile().getPath();
        }
        return path;
    }

    private void transferToExcel(String fileName) {
        File csvFile = new File(fileName);
        String folderPath = csvFile.getParent();
        String name = csvFile.getName();
        String subName = name.substring(0, name.lastIndexOf("."));
        String excelFileName = folderPath + "\\" + subName + ".xlsx";
        File excelFile = new File(excelFileName);

        System.out.println(excelFileName);
    }


}
