package com.utility;
import javax.swing.JOptionPane;

import javax.swing.JFrame;

public class AlertBox
{
public static void alert(String message)
{
 //Create a window using JFrame with title ( Message box appear in JFrame )
 JFrame frame=new JFrame("Message box appear in JFrame");

 //Set default close operation for JFrame
 frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

 //Set JFrame size
 frame.setSize(400,400);

 //Make JFrame locate at center on screen
 frame.setLocationRelativeTo(null);

 //Make JFrame visible
 //frame.setVisible(true);


 //Pop up a message box with text ( I am a message dialog ) in created JFrame
 JOptionPane.showMessageDialog(frame,message);
}

}
