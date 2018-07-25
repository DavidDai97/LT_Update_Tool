import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.util.*;
import java.util.concurrent.locks.LockSupport;

public class MainGUI {
    public static ArrayList<String> itemNumList = new ArrayList<>();
    public static Map<String, Item> itemMap = new HashMap<>();
    public static ArrayList<String> itemMasterArr = new ArrayList<>();
    public static ArrayList<String> poFileArr = new ArrayList<>();
    public static  JProgressBar processBar;

    public static void main(String[] args){
        getFileList("../Source");
        ExcelProcess.initializeFormat();
        createMainFrame();
    }

    private static void getFileList(String directory) {
        File f = new File(directory);
        File[] files = f.listFiles();
        if(files == null) return;
        for (int i = 0; i < files.length; i++) {
            if(files[i].getName().contains("Master")){
                itemMasterArr.add(files[i].getName());
            }
            else if(files[i].getName().contains("Receiving")){
                poFileArr.add(files[i].getName());
            }
        }
    }

    private static JFrame createFrame(int x, int y, int width, int height, java.awt.Color colourUse, String title, LayoutManager layoutUse){
        JFrame resultFrame = new JFrame(title);
        resultFrame.setBounds(x, y, width,height);
        resultFrame.setBackground(colourUse);
        resultFrame.setResizable(false);
        resultFrame.setLayout(layoutUse);
        resultFrame.setDefaultCloseOperation(WindowConstants.DO_NOTHING_ON_CLOSE);
        ((JPanel) resultFrame.getContentPane()).setBorder(BorderFactory.createEmptyBorder(25, 25, 25, 25));
        return resultFrame;
    }

    private static void createMainFrame(){
        JFrame mainFrame = createFrame(400, 100, 600, 350, Color.LIGHT_GRAY,
                "Lead Time Update Tool", new GridLayout(4, 1, 50, 10));
        JPanel labelPanel = new JPanel();
        labelPanel.setLayout(new GridLayout(1, 2, 50, 0));
        JLabel masterLabel = new JLabel("Item Master Choose:", JLabel.CENTER);
        masterLabel.setFont(new Font("Arial", Font.BOLD, 18));
        JLabel poReceiveLabel = new JLabel("PO Receiving Choose:", JLabel.CENTER);
        poReceiveLabel.setFont(new Font("Arial", Font.BOLD, 18));
        labelPanel.add(poReceiveLabel);
        labelPanel.add(masterLabel);
        JPanel scrollPanels = new JPanel();
        scrollPanels.setLayout(new GridLayout(1, 2, 50, 0));
        JList masterList = new JList(itemMasterArr.toArray(new String[0]));
        masterList.setVisibleRowCount(5);
        JList poList = new JList(poFileArr.toArray(new String[0]));
        poList.setVisibleRowCount(5);
        JScrollPane poScrollPane = new JScrollPane(poList);
        JScrollPane masterScrollPane = new JScrollPane(masterList);
        scrollPanels.add(poScrollPane);
        scrollPanels.add(masterScrollPane);
        JButton processButton = new JButton("Process");
        processButton.setFont(new Font("Arial", Font.BOLD, 30));
        processButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String masterFile = (String)masterList.getSelectedValue();
                String poFile = (String)poList.getSelectedValue();
                System.out.println(poFile + ": " + masterFile);
                MyRunnable processRunnable = new MyRunnable(poFile, masterFile);
                Thread processThread = new Thread(processRunnable);
                processThread.start();
            }
        });
        processBar = new JProgressBar();
        processBar.setStringPainted(true);
        processBar.setMinimum(0);
        processBar.setMaximum(100);
        processBar.setBackground(Color.GREEN);
        processBar.setValue(0);
        processBar.setBackground(Color.BLACK);
        processBar.setFont(new Font("Arial", Font.BOLD, 22));
        mainFrame.addWindowListener(new MyWin());
        mainFrame.add(labelPanel);
        mainFrame.add(scrollPanels);
        mainFrame.add(processButton);
        mainFrame.add(processBar);
        mainFrame.setVisible(true);
    }
}
