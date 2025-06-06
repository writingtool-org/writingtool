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
package org.writingtool.aisupport;

import java.util.ResourceBundle;

import org.writingtool.WtDocumentCache;
import org.writingtool.WtSingleDocument;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.WtDocumentsHandler.WaitDialogThread;
import org.writingtool.aisupport.WtAiRemote.AiCommand;
import org.writingtool.config.WtConfiguration;
import org.writingtool.dialogs.WtAiDialog;
import org.writingtool.dialogs.WtAiResultDialog;
import org.writingtool.dialogs.WtOptionPane;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtViewCursorTools;

import com.sun.star.frame.XController;
import com.sun.star.frame.XModel;
import com.sun.star.lang.Locale;
import com.sun.star.lang.XComponent;
import com.sun.star.text.XParagraphCursor;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextViewCursor;
import com.sun.star.text.XTextViewCursorSupplier;
import com.sun.star.uno.UnoRuntime;

/**
 * Class to execute changes of paragraphs by AI
 * @since 1.0
 * @author Fred Kruse
 */
public class WtAiParagraphChanging extends Thread {

  private static final ResourceBundle messages = WtOfficeTools.getMessageBundle();

  private int debugMode = WtOfficeTools.DEBUG_MODE_AI;   //  should be false except for testing

  public final static String WAIT_TITLE = messages.getString("loAiWaitDialogTitle");
  public final static String WAIT_MESSAGE = messages.getString("loAiWaitDialogMessage");

  private final WtSingleDocument document;
  private final WtConfiguration config;
  private final AiCommand commandId;
  
  private static WtAiDialog aiDialog = null;
  private WaitDialogThread waitDialog = null;
  
  public WtAiParagraphChanging(WtSingleDocument document, WtConfiguration config, AiCommand commandId) {
    this.document = document;
    this.config = config;
    this.commandId = commandId;
  }
  
  @Override
  public void run() {
    runAiChangeOnParagraph();
  }
  
  public static WtAiDialog getAiDialog() {
    return aiDialog;
  }
  
  public static void setCloseAiDialog() {
    aiDialog = null;
  }
  
  private void runAiChangeOnParagraph() {
    try {
      if (commandId == AiCommand.GeneralAi) {
        if (aiDialog == null) {
          waitDialog = new WaitDialogThread(WAIT_TITLE, WAIT_MESSAGE);
          aiDialog = new WtAiDialog(document, waitDialog, messages);
          aiDialog.start();
        } else {
          aiDialog.toFront();
        }
        return;
      }
      if (debugMode > 1) {
        WtMessageHandler.printToLogFile("AiParagraphChanging: runAiChangeOnParagraph: commandId: " + commandId);
      }
      XComponent xComponent = document.getXComponent();
      WtViewCursorTools viewCursor = new WtViewCursorTools(xComponent);
      TextParagraph tPara = viewCursor.getViewCursorParagraph();
      WtDocumentCache docCache = document.getDocumentCache();
      String text = docCache.getTextParagraph(tPara);
      Locale locale = docCache.getTextParagraphLocale(tPara);
//      String text = getViewCursorParagraph(xComponent);
      if (text == null || text.trim().isEmpty()) {
        return;
      }
      String instruction = null;
      boolean onlyPara = false;
      float temp = WtAiRemote.CORRECT_TEMPERATURE;
      if (commandId == AiCommand.CorrectGrammar) {
        instruction = WtAiRemote.getInstruction(WtAiRemote.CORRECT_INSTRUCTION, locale);
        onlyPara = true;
      } else if (commandId == AiCommand.ImproveStyle) {
        instruction = WtAiRemote.getInstruction(WtAiRemote.STYLE_INSTRUCTION, locale);
        onlyPara = true;
      } else if (commandId == AiCommand.ReformulateText) {
        instruction = WtAiRemote.getInstruction(WtAiRemote.REFORMULATE_INSTRUCTION, locale);
        onlyPara = true;
        temp = WtAiRemote.REFORMULATE_TEMPERATURE;
      } else if (commandId == AiCommand.ExpandText) {
        instruction = WtAiRemote.getInstruction(WtAiRemote.EXPAND_INSTRUCTION, locale);
        temp = WtAiRemote.EXPAND_TEMPERATURE;
      } else {
        instruction = WtOptionPane.showInputDialog(null, messages.getString("loMenuAiGeneralCommandMessage"), 
          messages.getString("loMenuAiGeneralCommandTitle"), WtOptionPane.QUESTION_MESSAGE);
        temp = WtAiRemote.EXPAND_TEMPERATURE;
      }
      waitDialog = new WaitDialogThread(WAIT_TITLE, WAIT_MESSAGE);
      waitDialog.start();
      WtAiRemote aiRemote = new WtAiRemote(document.getMultiDocumentsHandler(), config);
      if (debugMode > 1) {
        WtMessageHandler.printToLogFile("AiParagraphChanging: runInstruction: instruction: " + instruction + ", text: " + text);
      }
      String output = aiRemote.runInstruction(instruction, text, temp, 1, locale, onlyPara);
      if (debugMode > 1) {
        WtMessageHandler.printToLogFile("AiParagraphChanging: runAiChangeOnParagraph: output: " + output);
      }
      WtAiResultDialog resultDialog = new WtAiResultDialog(document, messages);
      if (output == null) {
        output = "";
      }
      resultDialog.setResult(output, tPara);
      resultDialog.start();
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    }
    if (waitDialog != null) {
      waitDialog.close();
      waitDialog = null;
    }
  }
  
  /** 
   * Returns the xText
   * Returns null if it fails
   */
  private static XText getXText(XComponent xComponent) {
    try {
      XTextDocument curDoc = UnoRuntime.queryInterface(XTextDocument.class, xComponent);
      if (curDoc == null) {
        return null;
      }
      return curDoc.getText();
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return null;           // Return null as method failed
    }
  }

  /** 
   * Returns ViewCursor 
   * Returns null if it fails
   */
  private static XTextViewCursor getViewCursor(XComponent xComponent) {
    try {
      XModel xModel = UnoRuntime.queryInterface(XModel.class, xComponent);
      if (xModel == null) {
        WtMessageHandler.printToLogFile("xModel == null");
        return null;
      }
      XController xController = xModel.getCurrentController();
      if (xController == null) {
        WtMessageHandler.printToLogFile("xController == null");
        return null;
      }
      XTextViewCursorSupplier xViewCursorSupplier =
          UnoRuntime.queryInterface(XTextViewCursorSupplier.class, xController);
      if (xViewCursorSupplier == null) {
        WtMessageHandler.printToLogFile("xViewCursorSupplier == null");
        return null;
      }
      return xViewCursorSupplier.getViewCursor();
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return null;           // Return null as method failed
    }
  }

  /** 
   * Returns Paragraph under ViewCursor 
   * Paragraph is selected
   *//*
  private static String getViewCursorParagraph(XComponent xComponent) {
    try {
      XTextViewCursor xVCursor = getViewCursor(xComponent);
      if (xVCursor == null) {
        MessageHandler.printToLogFile("xVCursor == null");
        return null;
      }
      XText xVCursorText = xVCursor.getText();
      XTextCursor xTCursor = xVCursorText.createTextCursorByRange(xVCursor.getStart());
      XParagraphCursor xPCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTCursor);
      if (xPCursor == null) {
        MessageHandler.showMessage("xPCursor == null");
        return null;
      }
      xPCursor.gotoStartOfParagraph(false);
      xPCursor.collapseToStart();
      xVCursor.gotoRange(xPCursor, false);
      xPCursor.gotoEndOfParagraph(false);
      xPCursor.collapseToEnd();
      xVCursor.gotoRange(xPCursor, true);
      return xVCursor.getString();
    } catch (Throwable t) {
      MessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return null;           // Return null as method failed
    }
  }
*/
  /** 
   * Inserts a Text to cursor position
   */
  public static void insertText(String text, XComponent xComponent, TextParagraph yPara, boolean override) {
    WtViewCursorTools vCursor = new WtViewCursorTools(xComponent);
    vCursor.setTextViewCursor(0, yPara);
    insertText(text, xComponent, override);
  }
  
  public static void insertText(String text, XComponent xComponent, boolean override) {
    if (text != null && xComponent != null) {
      try {
        XText xText = getXText(xComponent);
        if (xText == null) {
          return;
        }
        XTextViewCursor xVCursor = getViewCursor(xComponent);
        if (override) {
          XText xVCursorText = xVCursor.getText();
          XTextCursor xTCursor = xVCursorText.createTextCursorByRange(xVCursor.getStart());
          XParagraphCursor xPCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTCursor);
          if (xPCursor == null) {
            WtMessageHandler.showMessage("xPCursor == null");
            return;
          }
          xPCursor.gotoStartOfParagraph(false);
          xPCursor.gotoEndOfParagraph(true);
          xPCursor.setString("");
        }
        xVCursor.collapseToStart();
        xText.insertString(xVCursor, text, false);
      } catch (Throwable t) {
        WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      }
    }
  }

  /**
   * class to run a dialog in a separate thread
   * closing if lost focus
   *//*
  public class WaitDialogThread extends Thread {
    private final String dialogName;
    private final String text;
    private JDialog dialog = null;
    private boolean isCanceled = false;
    JProgressBar progressBar;

    public WaitDialogThread(String dialogName, String text) {
      this.dialogName = dialogName;
      this.text = text;
    }

    @Override
    public void run() {
      JLabel textLabel = new JLabel(text);
      JButton cancelBottom = new JButton("Abbrechen");
      cancelBottom.addActionListener(e -> {
        close_intern();
      });
      progressBar = new JProgressBar();
      progressBar.setIndeterminate(true);
      dialog = new JDialog();
      Container contentPane = dialog.getContentPane();
      dialog.setName("InformationThread");
      dialog.setTitle(dialogName);
      dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
      dialog.addWindowListener(new WindowListener() {
        @Override
        public void windowOpened(WindowEvent e) {
        }
        @Override
        public void windowClosing(WindowEvent e) {
          close_intern();
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
      JPanel panel = new JPanel();
      panel.setLayout(new GridBagLayout());
      GridBagConstraints cons = new GridBagConstraints();
      cons.insets = new Insets(16, 24, 16, 24);
      cons.gridx = 0;
      cons.gridy = 0;
      cons.weightx = 1.0f;
      cons.weighty = 10.0f;
      cons.anchor = GridBagConstraints.CENTER;
      cons.fill = GridBagConstraints.BOTH;
      panel.add(textLabel, cons);
      cons.gridy++;
      panel.add(progressBar, cons);
      cons.gridy++;
      cons.fill = GridBagConstraints.NONE;
      panel.add(cancelBottom, cons);
      contentPane.setLayout(new GridBagLayout());
      cons = new GridBagConstraints();
      cons.insets = new Insets(16, 32, 16, 32);
      cons.gridx = 0;
      cons.gridy = 0;
      cons.weightx = 1.0f;
      cons.weighty = 1.0f;
      cons.anchor = GridBagConstraints.NORTHWEST;
      cons.fill = GridBagConstraints.BOTH;
      contentPane.add(panel);
      dialog.pack();
      Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
      Dimension frameSize = dialog.getSize();
      dialog.setLocation(screenSize.width / 2 - frameSize.width / 2,
          screenSize.height / 2 - frameSize.height / 2);
      dialog.setAutoRequestFocus(true);
      dialog.setAlwaysOnTop(true);
      dialog.toFront();
//      if (debugMode > 1) {
//        MessageHandler.printToLogFile("WaitDialogThread: run: Dialog is running");
//      }
      dialog.setVisible(true);
      if (isCanceled) {
        dialog.setVisible(false);
        dialog.dispose();
      }
    }
    
    public boolean canceled() {
      return isCanceled;
    }
    
    public void close() {
      close_intern();
    }
    
    private void close_intern() {
//      if (debugMode > 1) {
//        MessageHandler.printToLogFile("WaitDialogThread: close: Dialog closed");
//      }
      isCanceled = true;
      if (dialog != null) {
        dialog.setVisible(false);
        dialog.dispose();
      }
    }
    
    public void initializeProgressBar(int min, int max) {
      progressBar.setMinimum(min);
      progressBar.setMaximum(max);
      progressBar.setStringPainted(true);
      progressBar.setIndeterminate(false);
    }
    
    public void setValueForProgressBar(int val) {
      progressBar.setValue(val);
    }
  }
*/
}
