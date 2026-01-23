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

import java.awt.Container;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Frame;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Image;
import java.awt.Insets;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
import java.util.ResourceBundle;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.ToolTipManager;

import org.writingtool.WtDocumentCache;
import org.writingtool.WtDocumentsHandler;
import org.writingtool.WtLanguageTool;
import org.writingtool.WtQuotesDetection;
import org.writingtool.WtSingleDocument;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtDocumentCursorTools;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtOfficeTools.DocumentType;

import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.Locale;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.text.XParagraphCursor;
import com.sun.star.uno.UnoRuntime;

import org.writingtool.tools.WtVersionInfo;

/**
 * Dialog to change paragraphs by AI
 * @since 1.1
 * @author Fred Kruse
 */
public class WtQuotesChangeDialog extends Thread implements ActionListener {
  
//  private final static int dialogWidth = 700;
//  private final static int dialogHeight = 150;

  private boolean debugMode = false;
  private boolean debugModeTm = false;
  
  private final ResourceBundle messages;
  private final JDialog dialog;
  private final Container contentPane;
  private final JButton close;
  private final JButton changeQuotes;
  private final JProgressBar checkProgress;
  private final Image ltImage;
  private final WtQuotesDetection quotesDetection;
  
  private final JLabel changeQuotesToLabel;
  private final JLabel numChangesLabel;
  private final JLabel numWrongQuotesLabel;
  private final JComboBox<String> quotes;
  
  private final JPanel mainPanel;

  private WtSingleDocument currentDocument;
  private DocumentType documentType;
  
  private int numWrongQuotes = 0;
  private int numChanges = 0;
  private int quoteIndex = 0;
  private int dialogX = -1;
  private int dialogY = -1;
  private boolean atWork = false;

  /**
   * the constructor of the class creates all elements of the dialog
   */
  public WtQuotesChangeDialog(WtSingleDocument document, ResourceBundle messages) {
    long startTime = 0;
    if (debugModeTm) {
      startTime = System.currentTimeMillis();
    }
    ltImage = WtOfficeTools.getWtImage();
    if (!WtDocumentsHandler.isJavaLookAndFeelSet()) {
      WtDocumentsHandler.setJavaLookAndFeel();
    }
    this.messages = messages;
    currentDocument = document;
    documentType = document.getDocumentType();
    
    dialog = new JDialog();
    contentPane = dialog.getContentPane();
    changeQuotesToLabel = new JLabel(messages.getString("loQuotesDialogChangeToLabel") + ":");
    quotes = new JComboBox<String>(getPossibleQuotes());
    numChangesLabel = new JLabel("(" + numChanges + " " + messages.getString("loQuotesNumberChangesLabel") + ")");
    numWrongQuotesLabel = new JLabel(numWrongQuotes + " " + messages.getString("loQuotesNumberWrongLabel"));

    changeQuotes = new JButton (messages.getString("loQuotesDialogChangeButton"));
    close = new JButton (messages.getString("loAiDialogCloseButton"));
    mainPanel = new JPanel();
    
    checkProgress = new JProgressBar();
    checkProgress.setStringPainted(true);
    checkProgress.setIndeterminate(false);
//    Locale locale = document.getDocumentCache().getDocumentLocale();
    quotesDetection = new WtQuotesDetection(document.getMultiDocumentsHandler().getContext());
    try {
      String txt = document.getDocumentCache().getDocAsString();
      numWrongQuotes = quotesDetection.numNotCorrectQuotes(txt, 0);
      numWrongQuotesLabel.setText(numWrongQuotes + " wrong quotes detected");
      if (debugMode) {
        WtMessageHandler.printToLogFile("CheckDialog: LtCheckDialog: LtCheckDialog called");
      }

      if (dialog == null) {
        WtMessageHandler.printToLogFile("CheckDialog: LtCheckDialog: LtCheckDialog == null");
      }
      String dialogName = messages.getString("loQuotesDialogTitle");
      dialog.setName(dialogName);
      dialog.setTitle(dialogName + " (" + WtVersionInfo.getWtNameWithInformation() + ")");
      dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
      ((Frame) dialog.getOwner()).setIconImage(ltImage);

      Font dialogFont = changeQuotesToLabel.getFont();
      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
          WtMessageHandler.printToLogFile("CheckDialog: Time to initialise Languages: " + runTime);
          startTime = System.currentTimeMillis();
      }

      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
          WtMessageHandler.printToLogFile("CheckDialog: Time to initialise suggestions, etc.: " + runTime);
          startTime = System.currentTimeMillis();
      }
      
      changeQuotes.setFont(dialogFont);
      changeQuotes.addActionListener(this);
      changeQuotes.setActionCommand("changeQuotes");
      changeQuotes.setEnabled(numWrongQuotes > 0);

      close.setFont(dialogFont);
      close.addActionListener(this);
      close.setActionCommand("close");

      dialog.addWindowListener(new WindowListener() {
        @Override
        public void windowOpened(WindowEvent e) {
        }
        @Override
        public void windowClosing(WindowEvent e) {
          closeDialog();
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
      
      quotes.addItemListener(e -> {
        if (e.getStateChange() == ItemEvent.SELECTED) {
          int index = quotes.getSelectedIndex();
          if (index != quoteIndex) {
            try {
              quoteIndex = index;
              numWrongQuotes = quotesDetection.numNotCorrectQuotes(document.getDocumentCache().getDocAsString(), quoteIndex);
              numWrongQuotesLabel.setText(numWrongQuotes + " wrong quotes detected");
              changeQuotes.setEnabled(numWrongQuotes > 0);
            } catch (Throwable t) {
              WtMessageHandler.showError(t);
            }
          }
        }
      });

      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
          WtMessageHandler.printToLogFile("CheckDialog: Time to initialise Buttons: " + runTime);
          startTime = System.currentTimeMillis();
      }
      
      //  Define Text panels

      //  Define left panel
      JPanel leftPanel1 = new JPanel();
      leftPanel1.setLayout(new GridBagLayout());
      GridBagConstraints cons11 = new GridBagConstraints();
      cons11.insets = new Insets(2, 0, 2, 0);
      cons11.gridx = 0;
      cons11.gridy = 0;
      cons11.anchor = GridBagConstraints.NORTHWEST;
      cons11.fill = GridBagConstraints.BOTH;
      cons11.weightx = 0.0f;
      cons11.weighty = 0.0f;
      cons11.gridy++;
      leftPanel1.add(changeQuotesToLabel, cons11);
      cons11.weightx = 10.0f;
      cons11.weighty = 10.0f;
      cons11.gridy++;
      leftPanel1.add(quotes, cons11);
      cons11.gridy++;
      leftPanel1.add(numChangesLabel, cons11);
      cons11.gridy++;
      leftPanel1.add(numWrongQuotesLabel, cons11);

      //  Define right panel
      JPanel rightPanel1 = new JPanel();
      rightPanel1.setLayout(new GridBagLayout());
      GridBagConstraints cons21 = new GridBagConstraints();
      cons21.insets = new Insets(2, 0, 2, 0);
      cons21.gridx = 0;
      cons21.gridy = 0;
      cons21.anchor = GridBagConstraints.SOUTHEAST;
      cons21.fill = GridBagConstraints.BOTH;
      cons21.weightx = 0.0f;
      cons21.weighty = 0.0f;
      cons21.gridy++;
      rightPanel1.add(changeQuotes, cons21);
      cons21.gridy++;
      rightPanel1.add(close, cons21);

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
      mainPanel.add(leftPanel1, cons1);
      cons1.gridx++;
      cons1.fill = GridBagConstraints.NONE;
      cons1.anchor = GridBagConstraints.SOUTHEAST;
      cons1.weightx = 1.0f;
      cons1.weighty = 1.0f;
      mainPanel.add(rightPanel1, cons1);

      //  Define check progress panel
      JPanel checkProgressPanel = new JPanel();
      checkProgressPanel.setLayout(new GridBagLayout());
      GridBagConstraints cons4 = new GridBagConstraints();
      cons4.insets = new Insets(4, 4, 4, 4);
      cons4.gridx = 0;
      cons4.gridy = 0;
      cons4.anchor = GridBagConstraints.NORTHWEST;
      cons4.fill = GridBagConstraints.HORIZONTAL;
      cons4.weightx = 4.0f;
      cons4.weighty = 0.0f;
      checkProgressPanel.add(checkProgress, cons4);

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
      cons.gridy++;
      cons.weighty = 0.0f;
      contentPane.add(checkProgressPanel, cons);

      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
//        if (runTime > OfficeTools.TIME_TOLERANCE) {
          WtMessageHandler.printToLogFile("CheckDialog: Time to initialise panels: " + runTime);
//        }
          startTime = System.currentTimeMillis();
      }

      dialog.pack();
      // center on screen:
      Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
//      Dimension frameSize = new Dimension(dialogWidth, dialogHeight);
//      dialog.setSize(frameSize);
      Dimension frameSize = dialog.getSize();
      frameSize = new Dimension((int)(frameSize.width * 1.2), (int)(frameSize.height * 1.2));
      dialog.setSize(frameSize);
      dialog.setLocation(screenSize.width / 2 - frameSize.width / 2,
          screenSize.height / 2 - frameSize.height / 2);
      dialog.setLocationByPlatform(true);
      
      ToolTipManager.sharedInstance().setDismissDelay(30000);
      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
//        if (runTime > OfficeTools.TIME_TOLERANCE) {
          WtMessageHandler.printToLogFile("CheckDialog: Time to initialise dialog size: " + runTime);
//        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      closeDialog();
    }
  }
  
  @Override
  public void run() {
    try {
      show();
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    }
  }

  /**
   * show the dialog
   * @throws Throwable 
   */
  public void show() throws Throwable {
    if (currentDocument == null || (currentDocument.getDocumentType() != DocumentType.WRITER 
          && currentDocument.getDocumentType() != DocumentType.IMPRESS)) {
      return;
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("WtQuotesChangeDialog: show: shoe dialog");
    }
    if (dialogX < 0 || dialogY < 0) {
      Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
      Dimension frameSize = dialog.getSize();
      dialogX = screenSize.width / 2 - frameSize.width / 2;
      dialogY = screenSize.height / 2 - frameSize.height / 2;
    }
    dialog.setLocation(dialogX, dialogY);
    dialog.setAutoRequestFocus(true);
    dialog.setVisible(true);
  }
  
  public void toFront() {
    dialog.setVisible(true);
    dialog.toFront();
  }

  /**
   * Actions of buttons
   */
  @Override
  public void actionPerformed(ActionEvent action) {
    if (!atWork) {
      try {
        if (debugMode) {
          WtMessageHandler.printToLogFile("WtQuotesChangeDialog: actionPerformed: Action: " + action.getActionCommand());
        }
        if (action.getActionCommand().equals("close")) {
          closeDialog();
        } else if (action.getActionCommand().equals("changeQuotes")) {
          changeQuotes();
        }
      } catch (Throwable e) {
        WtMessageHandler.showError(e);
        closeDialog();
      }
    }
  }
  
  private void replaceParagraph(TextParagraph textPara, String str, Locale locale, WtDocumentCursorTools docCursor) 
      throws UnknownPropertyException, PropertyVetoException, IllegalArgumentException, WrappedTargetException {
    XParagraphCursor pCursor = docCursor.getParagraphCursor(textPara);
    if (pCursor == null) {
      return;
    }
    pCursor.gotoStartOfParagraph(false);
    pCursor.gotoEndOfParagraph(true);
    pCursor.setString(str);
    XPropertySet xCursorProps = UnoRuntime.queryInterface(
        XPropertySet.class, pCursor );
    xCursorProps.setPropertyValue ( "CharLocale", locale);
  }
 
  private void writeText() {
    try {
      changeQuotes.setEnabled(false);
      quotes.setEnabled(false);
      checkProgress.setMinimum(0);
      checkProgress.setMaximum(numWrongQuotes);
      WtDocumentCursorTools docCursor = currentDocument.getDocumentCursorTools();
      WtDocumentCache docCache = currentDocument.getDocumentCache();
      Locale locale = docCache.getDocumentLocale();
      WtConfiguration conf = currentDocument.getMultiDocumentsHandler().getConfiguration();
      WtLanguageTool lt = null;
      boolean useQueue = conf.useTextLevelQueue() && !conf.noBackgroundCheck();
      if (useQueue) {
        lt = currentDocument.getMultiDocumentsHandler().getLanguageTool();
      }
      numChanges = 0;
      for(int i = 0; i < docCache.size(); i++) {
        String str = docCache.getFlatParagraph(i);
        String out = quotesDetection.changeToCorrectQuote(str, quoteIndex);
        TextParagraph textPara = docCache.getNumberOfTextParagraph(i);
        int nChanges = quotesDetection.getNumChanges();
        if (nChanges > 0) {
          numChanges += nChanges;
          replaceParagraph(textPara, out, locale, docCursor);
          docCache.setFlatParagraph(i, out);
          currentDocument.removeResultCache(i, true);
          currentDocument.removeIgnoredMatch(i, true);
          currentDocument.removePermanentIgnoredMatch(i, true);
          if (useQueue) {
            for (int j = 1; j < lt.getNumMinToCheckParas().size(); j++) {
              currentDocument.addQueueEntry(i, j, lt.getNumMinToCheckParas().get(j), currentDocument.getDocID(), true);
            }
          }
          setProgressValue(numChanges, true);
        }
      }
      numChangesLabel.setText("(" + numChanges + " " + messages.getString("loQuotesNumberChangesLabel") + ")");
      String txt = currentDocument.getDocumentCache().getDocAsString();
      numWrongQuotes = quotesDetection.numNotCorrectQuotes(txt, quoteIndex);
      numWrongQuotesLabel.setText(numWrongQuotes + " " + messages.getString("loQuotesNumberWrongLabel"));
      changeQuotes.setEnabled(numWrongQuotes > 0);
      quotes.setEnabled(true);
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }

  private void changeQuotes() throws Throwable {
    if (documentType == DocumentType.WRITER) {
      Thread t = new Thread(new Runnable() {
        public void run() {
          try {
            writeText();
          } catch (Throwable e) {
            WtMessageHandler.showError(e);
            closeDialog();
          }
        }
      });
      t.start();
    }
  }

  void setProgressValue(int value, boolean setText) {
    int max = checkProgress.getMaximum();
    int val = value < 0 ? 1 : value + 1;
    if (val > max) {
      val = max;
    }
    checkProgress.setValue(val);
    if (setText) {
      int p = (int) (((val * 100) / max) + 0.5);
      checkProgress.setString(p + " %  ( " + val + " / " + max + " )");
      checkProgress.setStringPainted(true);
    }
  }

  /**
   * returns an array of all possible quote pairs supported by WT
   */
  private String[] getPossibleQuotes() {
    int nQuotes = WtQuotesDetection.startSymbols.size();
    String quotes[] = new String[nQuotes];
    for (int i = 0; i < nQuotes; i++) {
      quotes[i] = WtQuotesDetection.startSymbols.get(i) + " " + WtQuotesDetection.endSymbols.get(i);
    }
    return quotes;
  }

  /**
   * closes the dialog
   */
  public void closeDialog() {
    dialog.setVisible(false);
//    if (debugMode) {
      WtMessageHandler.printToLogFile("WtQuotesChangeDialog: closeDialog: Close change quotes dialog");
//    }
    atWork = false;
    WtDocumentsHandler.closeChangeQuotesDialog();
  }
  

}
