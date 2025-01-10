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

import java.awt.Component;

import javax.swing.JOptionPane;

import org.writingtool.WtDocumentsHandler;
import org.writingtool.tools.WtGeneralTools;
import org.writingtool.tools.WtMessageHandler;

/**
 * simple panes for information and confirmation 
 * (adapter for JOptionPane)
 * NOTE: JOptionPane does not work with FlatLight and FlatDark themes
 * @since 1.2
 * @author Fred Kruse
 */
public class WtOptionPane {

  public static final int OK_OPTION = JOptionPane.OK_OPTION;
  public static final int CANCEL_OPTION = JOptionPane.CANCEL_OPTION;
  public static final int OK_CANCEL_OPTION = JOptionPane.OK_CANCEL_OPTION;
  public static final int QUESTION_MESSAGE = JOptionPane.QUESTION_MESSAGE;
  public static final int ERROR_MESSAGE = JOptionPane.ERROR_MESSAGE;
  public static final int INFORMATION_MESSAGE = JOptionPane.INFORMATION_MESSAGE;
  
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
      JOptionPane.showMessageDialog(parent, msg);
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
      JOptionPane.showMessageDialog(parent, msg, title, opt);
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
  
  public static String showInputDialog(Component parent, String msg, String initial) {
    int theme = WtDocumentsHandler.getJavaLookAndFeelSet();
    try {
      WtGeneralTools.setJavaLookAndFeel(WtGeneralTools.THEME_SYSTEM);
      String txt = JOptionPane.showInputDialog(parent, msg, initial);
      WtGeneralTools.setJavaLookAndFeel(theme);
      return txt;
    } catch (Exception e) {
      WtMessageHandler.printException(e);
    }
    return null;
  }
  
  public static int showConfirmDialog (Component parent, String msg, String title, int opt) {
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
/*
    if (showMessageBox(null, MessageBoxType.QUERYBOX, title, msg) == MessageBoxResults.OK) {
      return OK_OPTION;
    }
    return CANCEL_OPTION;
*/
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
  
  

}
