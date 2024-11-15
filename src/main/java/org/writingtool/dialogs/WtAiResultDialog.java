/* WritingTool, a LibreOffice Extension based on LanguageTool
 * Copyright (C) 2024 Fred Kruse (https://fk-es.de)
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
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
import java.util.ResourceBundle;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextPane;
import javax.swing.ToolTipManager;

import org.writingtool.WtDocumentsHandler;
import org.writingtool.WtSingleDocument;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.aisupport.WtAiParagraphChanging;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtOfficeTools.DocumentType;
import org.writingtool.tools.WtVersionInfo;

/**
 * Dialog to change paragraphs by AI
 * @since WT 1.1
 * @author Fred Kruse
 */
public class WtAiResultDialog extends Thread implements ActionListener {
  
  private final static int dialogWidth = 700;
  private final static int dialogHeight = 150;

  private boolean debugMode = false;
  private boolean debugModeTm = false;
  
  private final JDialog dialog;
  private final Container contentPane;
  private final JButton close;
//  private JProgressBar checkProgress;
  private final Image ltImage;
  
  private final JLabel resultLabel;
  private final JTextPane result;
  private final JButton overrideParagraph; 
  
  private final JPanel mainPanel;

  private WtSingleDocument currentDocument;
  private WtDocumentsHandler documents;
  private DocumentType documentType;
  private TextParagraph yPara;
  
  private int dialogX = -1;
  private int dialogY = -1;
  private boolean atWork = false;

  /**
   * the constructor of the class creates all elements of the dialog
   */
  public WtAiResultDialog(WtSingleDocument document, ResourceBundle messages) {
    documents = document.getMultiDocumentsHandler();
    long startTime = 0;
    if (debugModeTm) {
      startTime = System.currentTimeMillis();
    }
    ltImage = WtOfficeTools.getLtImage();
    if (!documents.isJavaLookAndFeelSet()) {
      WtDocumentsHandler.setJavaLookAndFeel();
    }
    
    currentDocument = document;
    documentType = document.getDocumentType();
    
    dialog = new JDialog();
    contentPane = dialog.getContentPane();
    resultLabel = new JLabel(messages.getString("loAiDialogResultLabel") + ":");
    result = new JTextPane();
    result.setBorder(BorderFactory.createLineBorder(Color.gray));
    overrideParagraph = new JButton (messages.getString("loAiDialogOverrideButton")); 
    close = new JButton (messages.getString("loAiDialogCloseButton"));
    mainPanel = new JPanel();
    
//    checkProgress.setStringPainted(true);
//    checkProgress.setIndeterminate(false);
    try {
      if (debugMode) {
        WtMessageHandler.printToLogFile("CheckDialog: LtCheckDialog: LtCheckDialog called");
      }

      if (dialog == null) {
        WtMessageHandler.printToLogFile("CheckDialog: LtCheckDialog: LtCheckDialog == null");
      }
      String dialogName = messages.getString("loAiResultDialogTitle");
      dialog.setName(dialogName);
      dialog.setTitle(dialogName + " (" + WtVersionInfo.getWtNameWithInformation() + ")");
      dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
      ((Frame) dialog.getOwner()).setIconImage(ltImage);

      Font dialogFont = resultLabel.getFont();
      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
          WtMessageHandler.printToLogFile("CheckDialog: Time to initialise Languages: " + runTime);
          startTime = System.currentTimeMillis();
      }

      JScrollPane resultPane = new JScrollPane(result);
      resultPane.setMinimumSize(new Dimension(100, 30));
      
      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
          WtMessageHandler.printToLogFile("CheckDialog: Time to initialise suggestions, etc.: " + runTime);
          startTime = System.currentTimeMillis();
      }
      
      overrideParagraph.setFont(dialogFont);
      overrideParagraph.addActionListener(this);
      overrideParagraph.setActionCommand("overrideParagraph");
      
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
      
      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
          WtMessageHandler.printToLogFile("CheckDialog: Time to initialise Buttons: " + runTime);
          startTime = System.currentTimeMillis();
      }
      
      //  Define Text panels

      //  Define 1. right panel
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
      rightPanel1.add(overrideParagraph, cons21);
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
      mainPanel.add(resultPane, cons1);
      cons1.gridx++;
      cons1.fill = GridBagConstraints.NONE;
      cons1.anchor = GridBagConstraints.SOUTHEAST;
      cons1.weightx = 1.0f;
      cons1.weighty = 1.0f;
      mainPanel.add(rightPanel1, cons1);
/*
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
//      checkProgressPanel.add(checkProgress, cons4);
*/
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
  //    cons.gridy++;
  //    cons.weighty = 0.0f;
  //    contentPane.add(checkProgressPanel, cons);

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
      Dimension frameSize = new Dimension(dialogWidth, dialogHeight);
      dialog.setSize(frameSize);
//      Dimension frameSize = dialog.getSize();
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
      WtMessageHandler.printToLogFile("CheckDialog: show: Goto next Error");
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
          WtMessageHandler.printToLogFile("CheckDialog: actionPerformed: Action: " + action.getActionCommand());
        }
        if (action.getActionCommand().equals("close")) {
          closeDialog();
        } else if (action.getActionCommand().equals("overrideParagraph")) {
          writeToParagraph(true);
        }
      } catch (Throwable e) {
        WtMessageHandler.showError(e);
        closeDialog();
      }
    }
  }
  
  private void writeToParagraph(boolean override) throws Throwable {
    if (documentType == DocumentType.WRITER) {
      WtAiParagraphChanging.insertText(result.getText(), currentDocument.getXComponent(), yPara, override);
    }
  }

  public void setResult(String text, TextParagraph yPara) {
    result.setText(text);
    this.yPara = yPara;
  }

  /**
   * closes the dialog
   */
  public void closeDialog() {
    dialog.setVisible(false);
    if (debugMode) {
      WtMessageHandler.printToLogFile("AiDialog: closeDialog: Close AI Dialog");
    }
    atWork = false;
  }
  

}
