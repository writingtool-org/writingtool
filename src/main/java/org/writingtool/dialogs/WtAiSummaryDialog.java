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
import java.awt.Container;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextPane;
import javax.swing.JTree;
import javax.swing.event.TreeSelectionEvent;
import javax.swing.event.TreeSelectionListener;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.TreePath;

import org.writingtool.WtChapter;
import org.writingtool.WtDocumentCache;
import org.writingtool.WtSingleDocument;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.WtDocumentsHandler;
import org.writingtool.WtDocumentsHandler.WaitDialogThread;
import org.writingtool.aisupport.WtAiParagraphChanging;
import org.writingtool.aisupport.WtAiRemote;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtOfficeTools.DocumentType;

import com.sun.star.lang.Locale;

/**
 * A dialog generates a summary for sub-chapters, chapters and the whole document
 * 
 * @author Fred Kruse
 */
public class WtAiSummaryDialog extends Thread implements ActionListener {
  
  private final static int MAX_SUMMARY_LENGTH = 1000;
  private final static String SUMMARY_INSTRUCTION = "Write a summary of the following text in less than " + MAX_SUMMARY_LENGTH + " characters";

  private final static int dialogWidth = 700;
  private final static int dialogHeight = 400;

  private boolean debugMode = WtOfficeTools.DEBUG_MODE_AI > 0;
  private boolean debugModeAiTm = WtOfficeTools.DEBUG_MODE_TA;

  private final ResourceBundle messages;
  private WtSingleDocument document;

  private final JDialog dialog = new JDialog();
  
  private final Container contentPane;
  private final JLabel chapterLabel;
  private final JLabel resultLabel;
  private final JTextPane result;
  private final JButton addSummary;
  private final JButton close;

  private JTree chapterTree;


  public WtAiSummaryDialog(WtSingleDocument document, ResourceBundle messages) {
    this.document = document;
    this.messages = messages;
    
    contentPane = dialog.getContentPane();
    chapterLabel = new JLabel(messages.getString("loAiDialogChapterLabel") + ":");
    resultLabel = new JLabel(messages.getString("loAiDialogResultLabelSummary") + ":");
    result = new JTextPane();
    result.setBorder(BorderFactory.createLineBorder(Color.gray));
    close = new JButton (messages.getString("loAiDialogCloseButton"));
    addSummary = new JButton (messages.getString("loAiDialogAddSummaryButton"));
  }
  
  @Override
  public void run() {
    try {
      show();
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    }
  }
  
  public void toFront() {
    dialog.setVisible(true);
    dialog.toFront();
  }

  public void close() {
    dialog.setVisible(false);
  }
  
  private void close_dialog() {
    document.getMultiDocumentsHandler().closeAiSummaryDialog();
  }
  
  public void show() {
    try {
      String summaryText = messages.getString("loAiDialogResultLabelSummary");
      
      dialog.setName(summaryText);
      dialog.setTitle(summaryText);
      dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
      
      dialog.addWindowListener(new WindowListener() {
        @Override
        public void windowOpened(WindowEvent e) {
        }
        @Override
        public void windowClosing(WindowEvent e) {
          close_dialog();
        }
        @Override
        public void windowClosed(WindowEvent e) {
        }
        @Override
        public void windowIconified(WindowEvent e) {
          close_dialog();
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

      Font dialogFont = resultLabel.getFont();

      close.setFont(dialogFont);
      close.addActionListener(this);
      close.setActionCommand("close");

      addSummary.setFont(dialogFont);
      addSummary.addActionListener(this);
      addSummary.setActionCommand("addSummary");

      
      
      WtChapter rootChapter = document.getDocumentCache().getChapters(messages.getString("loAiDialogDocument"));
      DefaultMutableTreeNode treeNode = createChapterTree(rootChapter);
      chapterTree = new JTree(treeNode);
      //  Add tree selection listener
      chapterTree.addTreeSelectionListener( new TreeSelectionListener() {
          public void valueChanged(TreeSelectionEvent event) {
             TreePath tp = event.getNewLeadSelectionPath();
             if (tp != null) {
               setEnableElements(false);
               ChapterNode node = (ChapterNode) tp.getLastPathComponent();
               CreateSummary summary = new CreateSummary(document, node.getChapter());
               summary.start();
             }
           }
         });
          
      
      //  Define Text panels

      //  Define left panel
      JPanel leftPanel = new JPanel();
      leftPanel.setLayout(new GridBagLayout());
      GridBagConstraints cons11 = new GridBagConstraints();
      cons11.insets = new Insets(4, 8, 4, 8);
      cons11.gridx = 0;
      cons11.gridy = 0;
      cons11.anchor = GridBagConstraints.NORTHWEST;
      cons11.fill = GridBagConstraints.BOTH;
      cons11.weightx = 0.0f;
      cons11.weighty = 0.0f;
      leftPanel.add(chapterLabel, cons11);
      cons11.gridy++;
      cons11.weightx = 10.0f;
      cons11.weighty = 10.0f;
      JScrollPane chapterPane = new JScrollPane(chapterTree);
      chapterPane.setMinimumSize(new Dimension(50, 50));
      leftPanel.add(chapterPane, cons11);

      //  Define right panel
      JPanel rightPanel = new JPanel();
      rightPanel.setLayout(new GridBagLayout());
      GridBagConstraints cons21 = new GridBagConstraints();
      cons21.insets = new Insets(4, 8, 4, 8);
      cons21.gridx = 0;
      cons21.gridy = 0;
      cons21.anchor = GridBagConstraints.NORTHWEST;
      cons21.fill = GridBagConstraints.BOTH;
      cons21.weightx = 0.0f;
      cons21.weighty = 0.0f;
      rightPanel.add(resultLabel, cons21);
      cons21.gridy++;
      cons21.weightx = 10.0f;
      cons21.weighty = 10.0f;
      JScrollPane resultPane = new JScrollPane(result);
      resultPane.setMinimumSize(new Dimension(150, 30));
      rightPanel.add(resultPane, cons21);
      cons21.gridy++;
      cons21.weightx = 0.0f;
      cons21.weighty = 0.0f;
      cons21.fill = GridBagConstraints.NONE;
      cons21.anchor = GridBagConstraints.SOUTHEAST;
      rightPanel.add(addSummary, cons21);

      //  Define main panel
      JPanel mainPanel = new JPanel();
      mainPanel.setLayout(new GridBagLayout());
      GridBagConstraints cons31 = new GridBagConstraints();
      cons31.insets = new Insets(4, 8, 4, 8);
      cons31.gridx = 0;
      cons31.gridy = 0;
      cons31.anchor = GridBagConstraints.NORTHWEST;
      cons31.fill = GridBagConstraints.BOTH;
      cons31.weightx = 10.0f;
      cons31.weighty = 10.0f;
      mainPanel.add(leftPanel, cons31);
      cons31.gridx++;
      mainPanel.add(rightPanel, cons31);

      contentPane.setLayout(new GridBagLayout());
      GridBagConstraints cons = new GridBagConstraints();
      cons.insets = new Insets(8, 8, 8, 8);
      cons.gridx = 0;
      cons.gridy = 0;
      cons.weightx = 10.0f;
      cons.weighty = 10.0f;
      cons.fill = GridBagConstraints.BOTH;
      cons.anchor = GridBagConstraints.NORTHWEST;
      contentPane.add(mainPanel, cons);
      cons.gridy++;
      cons.weightx = 0.0f;
      cons.weighty = 0.0f;
      cons.fill = GridBagConstraints.NONE;
      cons.anchor = GridBagConstraints.SOUTHEAST;
      contentPane.add(close, cons);
      
      dialog.pack();
      Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
      Dimension frameSize = new Dimension(dialogWidth, dialogHeight);
      dialog.setSize(frameSize);
//      Dimension frameSize = dialog.getSize();
      dialog.setLocation(screenSize.width / 2 - frameSize.width / 2,
          screenSize.height / 2 - frameSize.height / 2);
      dialog.setLocationByPlatform(true);
      dialog.setAutoRequestFocus(true);
      dialog.setVisible(true);
//      dialog.setAlwaysOnTop(true);
      dialog.toFront();
      
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }
  
  void setEnableElements(boolean enable) {
    chapterLabel.setEnabled(enable);
    resultLabel.setEnabled(enable);
    result.setEnabled(enable);
    chapterTree.setEnabled(enable);
    addSummary.setEnabled(enable);
  }
  
  private void createSubChapterNodes(ChapterNode parent) {
    WtChapter parentChapter = parent.getChapter();
    for (WtChapter childChapter : parentChapter.children) {
      ChapterNode child = new ChapterNode(childChapter);
      createSubChapterNodes(child);
      parent.add(child);
    }
  }
  
  private DefaultMutableTreeNode createChapterTree(WtChapter rootChapter) {
    ChapterNode root = new ChapterNode(rootChapter);
    createSubChapterNodes(root);
    return root;
  }

  /**
   * Actions of buttons
   */
  @Override
  public void actionPerformed(ActionEvent action) {
    try {
      if (debugMode) {
        WtMessageHandler.printToLogFile("CheckDialog: actionPerformed: Action: " + action.getActionCommand());
      }
      if (action.getActionCommand().equals("close")) {
        close_dialog();
      } else if (action.getActionCommand().equals("addSummary")) {
        writeToParagraph(false);
      }
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
      close_dialog();
    }
  }
  
  private void writeToParagraph(boolean override) throws Throwable {
    if (document.getDocumentType() == DocumentType.WRITER) {
      WtAiParagraphChanging.insertText(result.getText(), document.getXComponent(), override);
    }
  }

  public class CreateSummary extends Thread {
    
    private final WtSingleDocument document;
    private final WtChapter chapter;
    WaitDialogThread waitDialog;
    int nDone = 0;
    
    public CreateSummary(WtSingleDocument document, WtChapter chapter) {
      this.document = document;
      this.chapter = chapter;
    }
    
    @Override
    public void run() {
      try {
        waitDialog = WtDocumentsHandler.getWaitDialog(messages.getString("loAiWaitDialogTitle"), 
            messages.getString("loAiWaitDialogMessage"));
        int iter = getNumberOfChapters(chapter);
        waitDialog.initializeProgressBar(0, iter);
        waitDialog.start();
        result.setText(getChapterSummary(chapter));
        waitDialog.close();
        setEnableElements(true);
      } catch (Throwable e) {
        WtMessageHandler.showError(e);
      }
    }

    private List<String> getChapterText(int from) {
      WtDocumentCache docCache = document.getDocumentCache();
      TextParagraph tPara = new TextParagraph(WtDocumentCache.CURSOR_TYPE_TEXT, from);
      boolean useQueue = document.getMultiDocumentsHandler().getConfiguration().useTextLevelQueue();
      int startPos = docCache.getStartOfParaCheck(tPara, -1, false, useQueue, true);
      int endPos = docCache.getEndOfParaCheck(tPara, -1, false, useQueue, true);
      List<String> text = new ArrayList<>();
      for (int i = startPos; i < endPos; i++) {
        tPara = new TextParagraph(WtDocumentCache.CURSOR_TYPE_TEXT, i);
        text.add(docCache.getTextParagraph(tPara));
      }
      return text;
    }

    private String getChapterSummary(WtChapter chapter) {
      if (chapter.summary == null) {
        if (chapter.children == null || chapter.children.isEmpty()) {
          List<String> text = getChapterText(chapter.from);
          chapter.setSummary(getSummary(text));
        } else {
          List<String> text = new ArrayList<>();
          for (WtChapter subChapter : chapter.children) {
            if (subChapter.summary == null) {
              getChapterSummary(subChapter); 
            }
            text.add(subChapter.summary);
          }
          chapter.setSummary(getSummary(text));
        }
      }
      return chapter.summary;
    }
      
    private String getSummary(List<String> text) {
      if (text == null || text.isEmpty()) {
        return "";
      }
      String mergedText = text.get(0);
      List<String> summaries = new ArrayList<>();
      for (int i = 1; i < text.size(); i++) {
        if (mergedText.length() + text.get(i).length() > WtAiRemote.MAX_TEXT_LENGTH) {
          summaries.add(getSingleSummary(mergedText));
          mergedText = text.get(i);
        } else {
          mergedText += " " + text.get(i);
        }
      }
      summaries.add(getSingleSummary(mergedText));
      if (summaries.size() == 1) {
        return summaries.get(0);
      } else {
        return getSummary(summaries);
      }
    }

    private String getSingleSummary(String text) {
      try {
        if (waitDialog.canceled()) {
          return null;
        }
        waitDialog.setValueForProgressBar(nDone, true);
        nDone++;
        WtDocumentsHandler documents = document.getMultiDocumentsHandler();
        Locale locale = document.getDocumentCache().getDocumentLocale();
        String instruction = WtAiRemote.getInstruction(SUMMARY_INSTRUCTION, locale);
        WtAiRemote aiRemote = new WtAiRemote(documents, documents.getConfiguration());
        long startTime = 0;
        if (debugModeAiTm) {
          startTime = System.currentTimeMillis();
        }
        if (debugMode) {
          WtMessageHandler.printToLogFile("AiParagraphChanging: runInstruction: instruction: " + instruction + ", text: " 
              + (text.length() > 50 ? text.substring(0, 50)+ " ..." : text));
        }
        String output = aiRemote.runInstruction(instruction, text, 0.0f, 0, locale, false, false);
        if (debugMode) {
          WtMessageHandler.printToLogFile("AiParagraphChanging: runInstruction: output: " + output);
        }
        if (debugModeAiTm) {
          long runTime = System.currentTimeMillis() - startTime;
          WtMessageHandler.printToLogFile("AiErrorDetection: getMatchesByAiRule: Time to run AI detection rule for instruction " 
                                           + instruction + ", text: " + (text.length() > 50 ? text.substring(0, 50)
                                           + " ..." : text)+ ": " + runTime);
        }
        return output;
      } catch (Throwable t) {
        WtMessageHandler.showError(t);
      }
      return null;
    }
    
    private int getNumberOfChapters(WtChapter chapter) {
      int num = 0;
      if (chapter.children == null || chapter.children.isEmpty()) {
        List<String> text = getChapterText(chapter.from);
        int len = 0;
        for (String para : text) {
          len += para.length();
        }
        num = len / WtAiRemote.MAX_TEXT_LENGTH;
        if (num * WtAiRemote.MAX_TEXT_LENGTH < len) {
          num ++;
        }
        int ntmp = (num * MAX_SUMMARY_LENGTH) / WtAiRemote.MAX_TEXT_LENGTH;
        if (ntmp * WtAiRemote.MAX_TEXT_LENGTH < num * MAX_SUMMARY_LENGTH) {
          ntmp ++;
        }
        num += ntmp;
      } else {
        for (WtChapter subChapter : chapter.children) {
          if (subChapter.summary == null) {
            num += getNumberOfChapters(subChapter); 
          }
        }
        int ntmp = (chapter.children.size() * MAX_SUMMARY_LENGTH) / WtAiRemote.MAX_TEXT_LENGTH;
        if (ntmp * WtAiRemote.MAX_TEXT_LENGTH < chapter.children.size() * MAX_SUMMARY_LENGTH) {
          ntmp ++;
        }
        num += ntmp;
      }
      return num;
    }
      
  }

  public class ChapterNode extends DefaultMutableTreeNode {

    private static final long serialVersionUID = 1L;
    private WtChapter chapter;
    
    public ChapterNode(WtChapter chapter) {
      this.chapter = chapter;
    }
    
    public WtChapter getChapter() {
      return chapter;
    }
    
    @Override
    public String toString() {
      return chapter.name;
    }

    
  }
}
