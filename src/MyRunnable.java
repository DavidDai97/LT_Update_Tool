import javax.swing.*;
import java.awt.*;
import java.io.FileNotFoundException;

public class MyRunnable implements Runnable{
    private String poFileName;
    private String masterFileName;

    public MyRunnable(String poFileName, String masterFileName){
        this.poFileName = poFileName;
        this.masterFileName = masterFileName;
    }

    @Override
    public void run() {
        try {
            MainGUI.processBar.setBackground(Color.pink);
            MainGUI.processBar.setValue(0);
            ExcelProcess.processData(poFileName, masterFileName);
            MainGUI.processBar.setValue(100);
            MainGUI.processBar.setBackground(Color.BLACK);
            JOptionPane.showMessageDialog(null, "Process finished", "Progress",
                    JOptionPane.WARNING_MESSAGE);
        }
        catch (FileNotFoundException e){
            System.out.println("Error: " + e.toString());
            JOptionPane.showMessageDialog(null, "Err: File not found", "Error message", JOptionPane.ERROR_MESSAGE);
        }
        catch (Exception e){
            System.out.println(e.toString());
            JOptionPane.showMessageDialog(null, "Unknown Error", "Error message", JOptionPane.ERROR_MESSAGE);
        }
    }
}
