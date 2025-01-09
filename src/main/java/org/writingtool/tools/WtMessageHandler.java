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
package org.writingtool.tools;

import java.awt.event.WindowEvent;
import java.awt.event.WindowFocusListener;
import java.io.BufferedWriter;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.util.Date;

import javax.swing.JDialog;
import javax.swing.JOptionPane;
import javax.swing.UIManager;

import org.languagetool.tools.Tools;
import org.writingtool.WtDocumentsHandler;
import org.writingtool.dialogs.WtOptionPane;

import com.sun.star.uno.XComponentContext;

/**
 * Writes Messages to screen or log-file
 * @since 1.0
 * @author Fred Kruse, Marcin Miłkowski
 */
public class WtMessageHandler {
  
  private static final String logLineBreak = System.lineSeparator();  //  LineBreak in Log-File (MS-Windows compatible)
  
  private static boolean isOpen = false;
  private static boolean isInit = false;
  private static boolean isInitSp = false;
  
  private static boolean testMode;
  
  WtMessageHandler(XComponentContext xContext, boolean isSpellchecker) {
    initLogFile(xContext, isSpellchecker);
  }

  /**
   * Initialize log-file
   */
  private static void initLogFile(XComponentContext xContext, boolean isSpellchecker) {
    if ((!isSpellchecker && !isInit) || (isSpellchecker && !isInitSp)) {
      if (isSpellchecker) {
        isInitSp = true;
      } else {
        isInit = true;
      }
      try (OutputStream stream = new FileOutputStream(WtOfficeTools.getLogFilePath(xContext, isSpellchecker));
          OutputStreamWriter writer = new OutputStreamWriter(stream, StandardCharsets.UTF_8);
          BufferedWriter br = new BufferedWriter(writer)
          ) {
        Date date = new Date();
        writer.write("WritingTool log from " + date + logLineBreak + logLineBreak);
        writer.write("WritingTool " + WtVersionInfo.wtVersion + " (" + WtVersionInfo.wtBuildDate + ")" + logLineBreak);
        writer.write("LanguageTool " + WtVersionInfo.ltVersion() + " (" + WtVersionInfo.ltBuildDate() + ", " 
            + WtVersionInfo.ltShortGitId() + ")" + logLineBreak);
        writer.write("OS: " + System.getProperty("os.name") + " " 
            + System.getProperty("os.version") + " on " + System.getProperty("os.arch") + logLineBreak);
        if (WtVersionInfo.ooName != null) { 
          writer.write(WtVersionInfo.ooName + " " + WtVersionInfo.ooVersion + WtVersionInfo.ooExtension
              + " (" + WtVersionInfo.ooVendor +"), " + WtVersionInfo.ooLocale + logLineBreak);
        }
        writer.write(WtOfficeTools.getJavaInformation() + logLineBreak + logLineBreak);
        if (WtVersionInfo.ioEx != null) {
          writer.write(Tools.getFullStackTrace(WtVersionInfo.ioEx) + logLineBreak + logLineBreak);
        }
        if (WtVersionInfo.thEx != null) {
          writer.write(Tools.getFullStackTrace(WtVersionInfo.thEx) + logLineBreak + logLineBreak);
        }
      } catch (Throwable t) {
        showError(t);
      }
    }
  }
  
  /**
   * Initialize MessageHandler
   */
  public static void init(XComponentContext xContext, boolean isSpellchecker) {
    initLogFile(xContext, isSpellchecker);
  }

  /**
   * Show an error in a dialog
   */
  public static void showError(Throwable e, boolean fromSpellCheck) {
    if (fromSpellCheck) {
      printSpellException(e);
    } else {
      printException(e);
    }
    if (testMode) {
      throw new RuntimeException(e);
    }
    String msg = "An error has occurred in WritingTool "
        + WtVersionInfo.wtVersion + " (" + WtVersionInfo.wtBuildDate + "):\n" + e + "\nStacktrace:\n";
    msg += Tools.getFullStackTrace(e);
    String metaInfo = "OS: " + System.getProperty("os.name") + " on "
        + System.getProperty("os.arch") + ", Java version "
        + System.getProperty("java.version") + " from "
        + System.getProperty("java.vm.vendor");
    msg += metaInfo;
    DialogThread dt = new DialogThread(msg, true);
    e.printStackTrace();
    dt.start();
  }

  /**
   * Show a grammar error in a dialog
   */
  public static void showError(Throwable e) {
    showError(e, false);
  }
  
  /**
   * Show a spell error in a dialog
   */
  public static void showSpellError(Throwable e) {
    showError(e, true);
  }

  /**
   * Write to spell log-file
   */
  public static void printToLogFile(String str, boolean fromSpellCheck) {
    try (OutputStream stream = new FileOutputStream(WtOfficeTools.getLogFilePath(fromSpellCheck), true);
        OutputStreamWriter writer = new OutputStreamWriter(stream, StandardCharsets.UTF_8);
        BufferedWriter br = new BufferedWriter(writer)
        ) {
      writer.write(str + logLineBreak);
    } catch (Throwable t) {
      showError(t);
    }
  }

  /**
   * Write to spell log-file
   */
  public static void printToSpellLogFile(String str) {
    printToLogFile(str, true);
  }

  /** 
   * Prints Exception to spell log-file  
   */
  public static void printSpellException(Throwable t) {
   printToLogFile(Tools.getFullStackTrace(t), true);
  }

  /**
   * Write to grammar log-file
   */
  public static void printToLogFile(String str) {
    printToLogFile(str, false);
  }

  /** 
   * Prints Exception to grammar log-file  
   */
  public static void printException(Throwable t) {
   printToLogFile(Tools.getFullStackTrace(t), false);
  }

  /**
   * Will throw exception instead of showing errors as dialogs - use only for test cases.
   */
  public static void setTestMode(boolean mode) {
    testMode = mode;
  }

  /**
   * Shows a message in a dialog box
   * @param txt message to be shown
   */
  public static void showMessage(String txt) {
    showMessage(txt, true);
  }

  static void showMessage(String txt, boolean toLogFile) {
    if (toLogFile) {
      printToLogFile(txt);
    }
    DialogThread dt = new DialogThread(txt, false);
    dt.run();
  }
  
  public static void showFullStackMessage (String msg) {
    try {
      throw new RuntimeException(msg);
    } catch (Throwable t) {
      showError(t);
    }
  }

  /**
   * run an information message in a separate thread
   * closing if lost focus
   */
  public static void showClosingInformationDialog(String text) {
    ClosingInformationThread informationDialog = new ClosingInformationThread(text);
    informationDialog.start();
  }
  
  /**
   * class to run a dialog in a separate thread
   */
  private static class DialogThread extends Thread {
    private final String text;
    private boolean isException;

    DialogThread(String text, boolean isException) {
      if (text == null || text.isBlank()) {
        text = "Error empty text";
      }
      this.text = text;
      this.isException = isException;
    }

    @Override
    public void run() {
      if (isException) {
        if (!isOpen) {
          isOpen = true;
          WtOptionPane.showMessageDialog(null, text);
          isOpen = false;
        }
      } else {
        WtOptionPane.showMessageDialog(null, text);
      }
    }
  }
  
  /**
   * class to run a dialog in a separate thread
   * closing if lost focus
   */
  private static class ClosingInformationThread extends Thread {
    private final String text;
    JDialog dialog;

    ClosingInformationThread(String text) {
      this.text = text;
    }

    @Override
    public void run() {
      try {
        int theme = WtDocumentsHandler.getJavaLookAndFeelSet();
        WtGeneralTools.setJavaLookAndFeel(WtGeneralTools.THEME_SYSTEM);
        JOptionPane pane = new JOptionPane(text, JOptionPane.INFORMATION_MESSAGE);
        dialog = pane.createDialog(null, UIManager.getString("OptionPane.messageDialogTitle", null));
        dialog.setModal(false);
        dialog.setAutoRequestFocus(true);
        dialog.setAlwaysOnTop(true);
        dialog.addWindowFocusListener(new WindowFocusListener() {
          @Override
          public void windowGainedFocus(WindowEvent e) {
          }
          @Override
          public void windowLostFocus(WindowEvent e) {
            dialog.setVisible(false);
          }
        });
        dialog.setVisible(true);
        WtGeneralTools.setJavaLookAndFeel(theme);
      } catch (Exception e) {
        WtMessageHandler.printException(e);
      }
    }
  }
  
}
