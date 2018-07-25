import javax.swing.*;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

public class MyWin extends WindowAdapter {

    @Override
    public void windowClosing(WindowEvent e){
        System.out.println("Exit Program");
        int option = JOptionPane.showOptionDialog(null, "Are you sure to exit the program?","Exit",
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.ERROR_MESSAGE,null,null,null);
        System.out.println(option);
        if(option == 0) {
            System.exit(0);
        }
    }

    @Override
    public void windowActivated(WindowEvent e){
        System.out.println("Focused");
    }

    @Override
    public void windowOpened(WindowEvent e){
        System.out.println("Opened");
    }
}
