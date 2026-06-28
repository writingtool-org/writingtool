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

import javax.swing.JOptionPane;

import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;

import com.sun.star.uno.XComponentContext;

/**
 * simple panes for information and confirmation 
 * (adapter for JOptionPane)
 * NOTE: JOptionPane does not work with FlatLight and FlatDark themes
 * @since 1.2
 * @author Fred Kruse
 */
public class WtOptionPane {

  public static final int OK_OPTION = JOptionPane.OK_OPTION;
  public static final int CANCEL_OPTION = JOptionPane.NO_OPTION;
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

  public static void showMessageDialog (Component parent, String msg) {
    try {
      JOptionPane.showMessageDialog(parent, msg);
    } catch (Exception e) {
      WtMessageHandler.printException(e);
    }
  }
  
  public static void showCloseLoMessageDialog (String msg, XComponentContext xContext) {
    try {
      JOptionPane.showMessageDialog(null, msg);
      WtOfficeTools.closeLO(xContext);
    } catch (Exception e) {
      WtMessageHandler.printException(e);
    }
  }
  
  public static void showMessageDialog (Component parent, String msg, String title, int opt) {
    try {
      JOptionPane.showMessageDialog(parent, msg, title, opt);
    } catch (Exception e) {
      WtMessageHandler.printException(e);
    }
  }
  
  public static String showInputDialog(Component parent, Object msg, String title, int opt) {
    try {
      String txt = JOptionPane.showInputDialog(parent, msg, title, opt);
      return txt;
    } catch (Exception e) {
      WtMessageHandler.printException(e);
    }
    return null;
  }
  
  public static String showInputDialog(Component parent, String title, String initial) {
    try {
      String txt = JOptionPane.showInputDialog(parent, title, initial);
      return txt;
    } catch (Exception e) {
      WtMessageHandler.printException(e);
    }
    return null;
  }
  
  public static int showConfirmDialog (Component parent, String msg, String title, int opt) {
    try {
      int ret = JOptionPane.showConfirmDialog(parent, msg, title, opt);
      return ret;
    } catch (Exception e) {
      WtMessageHandler.printException(e);
    }
    return CANCEL_OPTION;
  }

}
