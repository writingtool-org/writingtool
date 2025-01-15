/* WritingTool, a LibreOffice Extension based on LanguageTool
 * Copyright (C) 2024 Fred Kruse (https://writingtool.org)
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301
 * USA
 */
package org.writingtool.dialogs;

import java.awt.Color;
import java.awt.Component;
import java.awt.Container;
import java.awt.Dialog;
import java.awt.Dimension;
import java.awt.Frame;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.HeadlessException;
import java.awt.Insets;
import java.awt.Toolkit;
import java.awt.Window;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
import java.util.ResourceBundle;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextField;
import javax.swing.UIManager;
import javax.swing.plaf.TextUI;

import org.writingtool.WtDocumentsHandler;
import org.writingtool.tools.WtGeneralTools;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;

/**
 * simple panes for information and confirmation 
 * (adapter for JOptionPane)
 * NOTE: JOptionPane does not work with FlatLight and FlatDark themes
 * @since 1.2
 * @author Fred Kruse
 */
public class WtOptionPane {

  private static final ResourceBundle messages = WtOfficeTools.getMessageBundle();

  public static final int OK_OPTION = JOptionPane.OK_OPTION;
  public static final int CANCEL_OPTION = JOptionPane.CANCEL_OPTION;
  public static final int OK_CANCEL_OPTION = JOptionPane.OK_CANCEL_OPTION;
  public static final int QUESTION_MESSAGE = JOptionPane.QUESTION_MESSAGE;
  public static final int ERROR_MESSAGE = JOptionPane.ERROR_MESSAGE;
  public static final int INFORMATION_MESSAGE = JOptionPane.INFORMATION_MESSAGE;
  public static final int MESSAGE_BOX = 0;
  public static final int INPUT_BOX = 1;
  public static final Color DARK_FOREGROUND = new Color(0xbabab2);
  public static final Color DARK_BACKGROUND = new Color(0x3c3f41);
  public static final Color DARK_BUTTON_BACKGROUND = new Color(0x4e5052);
  public static final Color DARK_BORDER = new Color(0x565c5f);

  private static int ret = CANCEL_OPTION;
  private static String out = null;


/*
  public static void showErrorDialog (Component parent, String msg) {
    try {
      showMessageBox(null, MessageBoxType.ERRORBOX, UIManager.getString("OptionPane.messageDialogTitle", null), msg);      
    } catch (Exception e) {
      WtMessageHandler.printException(e);
    }
  }
*/  
  public static void showMessageDialog (Component parent, String msg) {
    int theme = WtDocumentsHandler.getJavaLookAndFeelSet();
    try {
      WtGeneralTools.setJavaLookAndFeel(WtGeneralTools.THEME_SYSTEM);
      showMessageBox(parent, msg, UIManager.getString("OptionPane.messageDialogTitle", null), OK_OPTION, true);
//      JOptionPane.showMessageDialog(parent, msg);
      WtGeneralTools.setJavaLookAndFeel(theme);
//      showMessageBox(null, MessageBoxType.WARNINGBOX, UIManager.getString("OptionPane.messageDialogTitle", null), msg);      
    } catch (Exception e) {
      WtMessageHandler.printException(e);
    }
  }
  
  public static void showMessageDialog (Component parent, String msg, String title, int opt) {
      
    int theme = WtDocumentsHandler.getJavaLookAndFeelSet();
    try {
      WtGeneralTools.setJavaLookAndFeel(WtGeneralTools.THEME_SYSTEM);
      showMessageBox(parent, msg, title, opt, true);
//      JOptionPane.showMessageDialog(parent, msg, title, opt);
      WtGeneralTools.setJavaLookAndFeel(theme);
    } catch (Exception e) {
      WtMessageHandler.printException(e);
    }
//    showMessageBox(null, MessageBoxType.WARNINGBOX, title, msg);      
  }
  
  public static String showInputDialog(Component parent, Object msg, String title, int opt) {
    int theme = WtDocumentsHandler.getJavaLookAndFeelSet();
    try {
      WtGeneralTools.setJavaLookAndFeel(WtGeneralTools.THEME_SYSTEM);
      String txt = JOptionPane.showInputDialog(parent, msg, title, opt);
      WtGeneralTools.setJavaLookAndFeel(theme);
      return txt;
    } catch (Exception e) {
      WtMessageHandler.printException(e);
    }
    return null;
  }
  
  public static String showInputDialog(Component parent, String title, String initial) {
/*    int theme = WtDocumentsHandler.getJavaLookAndFeelSet();
    try {
      WtGeneralTools.setJavaLookAndFeel(WtGeneralTools.THEME_SYSTEM);
      String txt = JOptionPane.showInputDialog(parent, msg, initial);
      WtGeneralTools.setJavaLookAndFeel(theme);
      return txt;
    } catch (Exception e) {
      WtMessageHandler.printException(e);
    }
    return null;
*/
    return showInputBox(parent, title, initial);
  }
  
  public static int showConfirmDialog (Component parent, String msg, String title, int opt) {
/*
    int theme = WtDocumentsHandler.getJavaLookAndFeelSet();
    try {
      WtGeneralTools.setJavaLookAndFeel(WtGeneralTools.THEME_SYSTEM);
      int ret = JOptionPane.showConfirmDialog(parent, msg, title, opt);
      WtGeneralTools.setJavaLookAndFeel(theme);
      return ret;
    } catch (Exception e) {
      WtMessageHandler.printException(e);
    }
    return CANCEL_OPTION;
*/
    return showMessageBox(parent, msg, title, opt, false);
  }
/*
 * This is commented out for further development
 * LO dialogs are not so flexible as java dialogs  

  public static short showMessageBox(XDialog dialog, MessageBoxType type, String sTitle, String sMessage) {
    XToolkit xToolkit;
    XComponentContext context = WtDocumentsHandler.getComponentContext();
    try {
      xToolkit = UnoRuntime.queryInterface(XToolkit.class,
            context.getServiceManager().createInstanceWithContext("com.sun.star.awt.Toolkit", context));
    } catch (Exception e) {
      return -1;
    }
    XMessageBoxFactory xMessageBoxFactory = UnoRuntime.queryInterface(XMessageBoxFactory.class, xToolkit);
    XWindowPeer xParentWindowPeer;
    if (dialog == null) {
      XWindow window = WtOfficeTools.getCurrentWindow(context);
      if (window == null) {
        return -1;
      }
      xParentWindowPeer = UnoRuntime.queryInterface(XWindowPeer.class, window);
    } else {
      xParentWindowPeer = UnoRuntime.queryInterface(XWindowPeer.class, dialog);
    }
    XMessageBox xMessageBox;
    if (type == MessageBoxType.QUERYBOX) {
      xMessageBox = xMessageBoxFactory.createMessageBox(xParentWindowPeer, type, MessageBoxButtons.BUTTONS_OK_CANCEL, sTitle, sMessage);
    } else {
      xMessageBox = xMessageBoxFactory.createMessageBox(xParentWindowPeer, type, MessageBoxButtons.BUTTONS_OK, sTitle, sMessage);
    }
    if (xMessageBox == null) {
      return -1;
    }
    return xMessageBox.execute();
  }
*/
  
  public static int showMessageBox(Component parent, String msg, String title, int opt, boolean isSystem) {
    try {
      ret = CANCEL_OPTION;
      int theme = WtDocumentsHandler.getJavaLookAndFeelSet();
//      WtGeneralTools.setJavaLookAndFeel(WtGeneralTools.THEME_SYSTEM);
      JDialog dialog;
      if (parent == null) {
        dialog = new JDialog();
      } else {
        Window parentWindow = getWindowForComponent(parent);
        if (parentWindow instanceof Frame) {
          dialog = new JDialog((Frame) parentWindow);
        } else {
          dialog = new JDialog((Dialog) parentWindow);
        }
      }
      dialog.setModal(true);
      Container contentPane = dialog.getContentPane();
      
      JLabel text = new JLabel();
      JButton cancelButton = new JButton (WtGeneralTools.getLabel(messages.getString("guiCancelButton"))); 
      JButton okButton = new JButton (WtGeneralTools.getLabel(messages.getString("guiOKButton")));
      JPanel mainPanel = new JPanel();
      msg = "<html><body>" + msg.replace("\n", "<br>") + "</body></html>";
      text.setText(msg);
      JScrollPane textPane = new JScrollPane(text);
      if (title == null) {
        title = "";
      }
      dialog.setTitle(title);
      dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);

      okButton.addActionListener(e -> {
        ret = OK_OPTION;
        dialog.setVisible(false);
      });
      okButton.setVisible(opt != CANCEL_OPTION);
      
      cancelButton.addActionListener(e -> {
        dialog.setVisible(false);
      });
      cancelButton.setVisible(opt != OK_OPTION);

      dialog.addWindowListener(new WindowListener() {
        @Override
        public void windowOpened(WindowEvent e) {
        }
        @Override
        public void windowClosing(WindowEvent e) {
          dialog.setVisible(false);
        }
        @Override
        public void windowClosed(WindowEvent e) {
        }
        @Override
        public void windowIconified(WindowEvent e) {
        }
        @Override
        public void windowDeiconified(WindowEvent e) {
        }
        @Override
        public void windowActivated(WindowEvent e) {
        }
        @Override
        public void windowDeactivated(WindowEvent e) {
        }
      });
      
      //  Define Button panel
      JPanel buttonPanel = new JPanel();
      buttonPanel.setLayout(new GridBagLayout());
      GridBagConstraints cons21 = new GridBagConstraints();
      cons21.insets = new Insets(2, 0, 2, 0);
      cons21.gridx = 0;
      cons21.gridy = 0;
      cons21.anchor = GridBagConstraints.SOUTHEAST;
      cons21.fill = GridBagConstraints.BOTH;
      cons21.weightx = 0.0f;
      cons21.weighty = 0.0f;
      buttonPanel.add(okButton, cons21);
      cons21.gridx++;
      buttonPanel.add(cancelButton, cons21);

      //  Define main panel
      mainPanel.setLayout(new GridBagLayout());
      GridBagConstraints cons1 = new GridBagConstraints();
      cons1.insets = new Insets(4, 4, 4, 4);
      cons1.gridx = 0;
      cons1.gridy = 0;
      cons1.anchor = GridBagConstraints.NORTHWEST;
      cons1.fill = GridBagConstraints.BOTH;
      cons1.weightx = 10.0f;
      cons1.weighty = 10.0f;
      mainPanel.add(textPane, cons1);
      cons1.insets = new Insets(14, 4, 4, 4);
      cons1.gridy++;
      cons1.fill = GridBagConstraints.BOTH;
      cons1.anchor = GridBagConstraints.CENTER;
      cons1.weightx = 1.0f;
      cons1.weighty = 1.0f;
      mainPanel.add(buttonPanel, cons1);

      contentPane.setLayout(new GridBagLayout());
      GridBagConstraints cons = new GridBagConstraints();
      cons.insets = new Insets(8, 8, 8, 8);
      cons.gridx = 0;
      cons.gridy = 0;
      cons.anchor = GridBagConstraints.NORTHWEST;
      cons.fill = GridBagConstraints.BOTH;
      cons.weightx = 1.0f;
      cons.weighty = 1.0f;
      contentPane.add(mainPanel, cons);
      
      if (isSystem && theme == WtGeneralTools.THEME_FLATDARK) {
        dialog.setBackground(DARK_BACKGROUND);
        contentPane.setBackground(DARK_BACKGROUND);
        mainPanel.setBackground(DARK_BACKGROUND);
        buttonPanel.setBackground(DARK_BACKGROUND);
        textPane.setBackground(DARK_BACKGROUND);
        text.setBackground(DARK_BACKGROUND);
        text.setForeground(DARK_FOREGROUND);
        text.setOpaque(true);
        okButton.setBackground(DARK_BUTTON_BACKGROUND);
        okButton.setForeground(DARK_FOREGROUND);
        okButton.setContentAreaFilled(false);
        okButton.setOpaque(true);
        okButton.setBorder(BorderFactory.createLineBorder(DARK_BORDER));
        cancelButton.setBackground(DARK_BUTTON_BACKGROUND);
        cancelButton.setForeground(DARK_FOREGROUND);
        cancelButton.setContentAreaFilled(false);
        cancelButton.setOpaque(true);
        cancelButton.setBorder(BorderFactory.createLineBorder(DARK_BORDER));
      }
      text.setBorder(BorderFactory.createLineBorder(text.getBackground()));

      dialog.pack();
      // center on screen:
      Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
      Dimension frameSize = dialog.getSize();
      dialog.setLocation(screenSize.width / 2 - frameSize.width / 2,
          screenSize.height / 2 - frameSize.height / 2);
      dialog.setLocationByPlatform(true);
      dialog.setVisible(true);
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return ret;
  }
  
  public static String showInputBox(Component parent, String title, String initial) {
    out = null;
    try {
      JDialog dialog;
      if (parent == null) {
        dialog = new JDialog();
      } else {
        Window parentWindow = getWindowForComponent(parent);
        if (parentWindow instanceof Frame) {
          dialog = new JDialog((Frame) parentWindow);
        } else {
          dialog = new JDialog((Dialog) parentWindow);
        }
      }
      dialog.setModal(true);
      Container contentPane = dialog.getContentPane();
      
      JTextField text = new JTextField();
      TextUI ui = text.getUI();
      if (ui == null) {
        WtMessageHandler.showMessage("showInputBox: UI == null");
      }
      JButton cancelButton = new JButton (WtGeneralTools.getLabel(messages.getString("guiCancelButton"))); 
      JButton okButton = new JButton (WtGeneralTools.getLabel(messages.getString("guiOKButton")));
      JPanel mainPanel = new JPanel();
      text.setEditable(true);
      text.setMinimumSize(new Dimension(100, 30));
      text.setBorder(BorderFactory.createLineBorder(Color.gray));
      text.updateUI();
      if (title == null) {
        title = "";
      }
      dialog.setTitle(title);
      dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);

      okButton.addActionListener(e -> {
        out = text.getText();
        dialog.setVisible(false);
      });
      
      cancelButton.addActionListener(e -> {
        out = null;
        dialog.setVisible(false);
      });

      dialog.addWindowListener(new WindowListener() {
        @Override
        public void windowOpened(WindowEvent e) {
        }
        @Override
        public void windowClosing(WindowEvent e) {
          out = null;
          dialog.setVisible(false);
        }
        @Override
        public void windowClosed(WindowEvent e) {
        }
        @Override
        public void windowIconified(WindowEvent e) {
        }
        @Override
        public void windowDeiconified(WindowEvent e) {
        }
        @Override
        public void windowActivated(WindowEvent e) {
        }
        @Override
        public void windowDeactivated(WindowEvent e) {
        }
      });
      
      //  Define button panel
      JPanel buttonPanel = new JPanel();
      buttonPanel.setLayout(new GridBagLayout());
      GridBagConstraints cons21 = new GridBagConstraints();
      cons21.insets = new Insets(2, 0, 2, 0);
      cons21.gridx = 0;
      cons21.gridy = 0;
      cons21.anchor = GridBagConstraints.SOUTHEAST;
      cons21.fill = GridBagConstraints.BOTH;
      cons21.weightx = 0.0f;
      cons21.weighty = 0.0f;
      buttonPanel.add(okButton, cons21);
      cons21.gridx++;
      buttonPanel.add(cancelButton, cons21);

      //  Define main panel
      mainPanel.setLayout(new GridBagLayout());
      GridBagConstraints cons1 = new GridBagConstraints();
      cons1.insets = new Insets(4, 4, 4, 4);
      cons1.gridx = 0;
      cons1.gridy = 0;
      cons1.anchor = GridBagConstraints.NORTHWEST;
      cons1.fill = GridBagConstraints.BOTH;
      cons1.weightx = 10.0f;
      cons1.weighty = 10.0f;
      mainPanel.add(text, cons1);
      cons1.insets = new Insets(14, 4, 4, 4);
      cons1.gridy++;
      cons1.fill = GridBagConstraints.BOTH;
      cons1.anchor = GridBagConstraints.CENTER;
      cons1.weightx = 1.0f;
      cons1.weighty = 1.0f;
      mainPanel.add(buttonPanel, cons1);

      contentPane.setLayout(new GridBagLayout());
      GridBagConstraints cons = new GridBagConstraints();
      cons.insets = new Insets(8, 8, 8, 8);
      cons.gridx = 0;
      cons.gridy = 0;
      cons.anchor = GridBagConstraints.NORTHWEST;
      cons.fill = GridBagConstraints.BOTH;
      cons.weightx = 1.0f;
      cons.weighty = 1.0f;
      contentPane.add(mainPanel, cons);
      
      dialog.pack();

      // center on screen:
      Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
      Dimension frameSize = dialog.getSize();
      dialog.setLocation(screenSize.width / 2 - frameSize.width / 2,
          screenSize.height / 2 - frameSize.height / 2);
      dialog.setLocationByPlatform(true);
      if (initial != null) {
        text.setText(initial);
      }
      dialog.setVisible(true);
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return out;
  }
  
  static Window getWindowForComponent(Component parentComponent) throws HeadlessException {
      if (parentComponent == null) {
        return null;
      }
      if (parentComponent instanceof Frame || parentComponent instanceof Dialog) {
        return (Window)parentComponent;
      }
      return getWindowForComponent(parentComponent.getParent());
  }
  
  

}
