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
import java.awt.Dialog.ModalityType;
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
import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JSlider;
import javax.swing.ToolTipManager;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import org.languagetool.Language;
import org.languagetool.Languages;
import org.writingtool.WtDocumentsHandler;
import org.writingtool.WtLinguisticServices;
import org.writingtool.WtSingleDocument;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtOfficeTools.DocumentType;
import org.writingtool.tools.WtVersionInfo;

import com.sun.star.lang.Locale;

/**
 * Dialog to change paragraphs by AI
 * @since 1.1
 * @author Fred Kruse
 */
public class WtAiTranslationDialog implements ActionListener {
  
  private final static int dialogWidth = 700;
  private final static int dialogHeight = 200;
  
  private final static float DEFAULT_TEMPERATURE = 0.0f;

  private boolean debugMode = false;
  private boolean debugModeTm = false;
  
  private final ResourceBundle messages;
  private final JDialog dialog;
  private final Container contentPane;
  private final Image ltImage;
  
  private final JLabel languageLabel;
  private final JComboBox<String> language;
  private final JLabel temperatureLabel;
  private final JSlider temperatureSlider;

  private final JButton translate; 
  private final JButton cancel;
  
  private final JPanel mainPanel;

  private WtSingleDocument currentDocument;
  private String startLang;
  private String selectedLang;
  private Locale locale = null;
  private float temperature = DEFAULT_TEMPERATURE;
  private int dialogX = -1;
  private int dialogY = -1;

  /**
   * the constructor of the class creates all elements of the dialog
   */
  public WtAiTranslationDialog(WtSingleDocument document, ResourceBundle messages) {
    this.messages = messages;
    long startTime = 0;
    if (debugModeTm) {
      startTime = System.currentTimeMillis();
    }
    ltImage = WtOfficeTools.getWtImage();
    if (!WtDocumentsHandler.isJavaLookAndFeelSet()) {
      WtDocumentsHandler.setJavaLookAndFeel();
    }
    
    currentDocument = document;
    
    dialog = new JDialog();
    contentPane = dialog.getContentPane();
    languageLabel = new JLabel(messages.getString("loAiDialogLanguageLabel") + ":");
    language = new JComboBox<String>(getPossibleLanguages());
    temperatureLabel = new JLabel(messages.getString("loAiDialogTranslateFreedom") + ":");
    temperatureSlider = new JSlider(0, 100, (int)(DEFAULT_TEMPERATURE*100));
    translate = new JButton (messages.getString("loAiDialogTranslateButton")); 
    cancel = new JButton (messages.getString("guiCancelButton"));
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
      String dialogName = messages.getString("loAiTranslationDialogTitle");
      dialog.setName(dialogName);
//      dialog.setTitle(dialogName);
      dialog.setTitle(dialogName + " (" + WtVersionInfo.getWtNameWithInformation() + ")");
      dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
      ((Frame) dialog.getOwner()).setIconImage(ltImage);

      Font dialogFont = languageLabel.getFont();
      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
          WtMessageHandler.printToLogFile("CheckDialog: Time to initialise Languages: " + runTime);
          startTime = System.currentTimeMillis();
      }

      language.setFont(dialogFont);
      language.setSelectedItem(startLang);
      locale = getLocaleFromLanguageName(startLang);
      selectedLang = startLang;
//      language.setToolTipText(formatToolTipText(languageHelp));
      language.addItemListener(e -> {
        if (e.getStateChange() == ItemEvent.SELECTED) {
          selectedLang = (String) language.getSelectedItem();
        }
      });
      
      temperatureLabel.setFont(dialogFont);
      temperatureSlider.setMajorTickSpacing(10);
      temperatureSlider.setMinorTickSpacing(5);
      temperatureSlider.setPaintTicks(true);
      temperatureSlider.setSnapToTicks(true);
      temperatureSlider.addChangeListener(new ChangeListener( ) {
        @Override
        public void stateChanged(ChangeEvent e) {
          int value = temperatureSlider.getValue();
          temperature = (float) (value / 100.);
        }
      });



      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
          WtMessageHandler.printToLogFile("CheckDialog: Time to initialise suggestions, etc.: " + runTime);
          startTime = System.currentTimeMillis();
      }
      
      translate.setFont(dialogFont);
      translate.addActionListener(this);
      translate.setActionCommand("translate");
      
      cancel.setFont(dialogFont);
      cancel.addActionListener(this);
      cancel.setActionCommand("cancel");

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

      //  Define 1. left panel
      JPanel leftPanel1 = new JPanel();
      leftPanel1.setLayout(new GridBagLayout());
      GridBagConstraints cons11 = new GridBagConstraints();
      cons11.insets = new Insets(2, 0, 2, 0);
      cons11.gridx = 0;
      cons11.gridy = 0;
      cons11.anchor = GridBagConstraints.NORTHWEST;
      cons11.fill = GridBagConstraints.NONE;
      cons11.weightx = 0.0f;
      cons11.weighty = 0.0f;
      leftPanel1.add(languageLabel, cons11);
      cons11.gridy++;
      cons11.fill = GridBagConstraints.HORIZONTAL;
      cons11.weightx = 10.0f;
      leftPanel1.add(language, cons11);
      cons11.gridy++;
      cons11.fill = GridBagConstraints.NONE;
      cons11.weightx = 0.0f;
      cons11.insets = new Insets(18, 0, 2, 0);
      leftPanel1.add(temperatureLabel, cons11);
      cons11.gridy++;
      cons11.insets = new Insets(2, 0, 2, 0);
      cons11.fill = GridBagConstraints.HORIZONTAL;
      cons11.weightx = 10.0f;
      leftPanel1.add(temperatureSlider, cons11);

      //  Define 1. right panel
      JPanel rightPanel1 = new JPanel();
      rightPanel1.setLayout(new GridBagLayout());
      GridBagConstraints cons21 = new GridBagConstraints();
      cons21.insets = new Insets(2, 0, 2, 0);
      cons21.gridx = 0;
      cons21.gridy = 0;
      cons21.anchor = GridBagConstraints.SOUTHEAST;
      cons21.fill = GridBagConstraints.HORIZONTAL;
      cons21.weightx = 0.0f;
      cons21.weighty = 0.0f;
      cons21.gridy++;
      rightPanel1.add(translate, cons21);
      cons21.gridy++;
      rightPanel1.add(cancel, cons21);

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

      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
        if (runTime > WtOfficeTools.TIME_TOLERANCE) {
          WtMessageHandler.printToLogFile("CheckDialog: Time to initialise panels: " + runTime);
        }
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
        if (runTime > WtOfficeTools.TIME_TOLERANCE) {
          WtMessageHandler.printToLogFile("CheckDialog: Time to initialise dialog size: " + runTime);
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      closeDialog();
    }
  }
  
  /**
   * run the dialog
   */
  public TranslationOptions run() {
    if (currentDocument == null || (currentDocument.getDocumentType() != DocumentType.WRITER 
          && currentDocument.getDocumentType() != DocumentType.IMPRESS)) {
      return null;
    }
    if (dialogX < 0 || dialogY < 0) {
      Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
      Dimension frameSize = dialog.getSize();
      dialogX = screenSize.width / 2 - frameSize.width / 2;
      dialogY = screenSize.height / 2 - frameSize.height / 2;
    }
    dialog.setLocation(dialogX, dialogY);
    dialog.setModalityType(ModalityType.DOCUMENT_MODAL);
    dialog.setAutoRequestFocus(true);
    dialog.setVisible(true);
    if(locale != null) {
      return new TranslationOptions(locale, temperature);
    }
    return null;
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
    try {
      if (debugMode) {
        WtMessageHandler.printToLogFile("CheckDialog: actionPerformed: Action: " + action.getActionCommand());
      }
      if (action.getActionCommand().equals("cancel")) {
        closeDialog();
      } else if (action.getActionCommand().equals("translate")) {
        translate();
      }
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
      closeDialog();
    }
  }
  
  /**
   * Start translation
   */
  
  private void translate() {
    if (selectedLang != null) {
      locale = getLocaleFromLanguageName(selectedLang);
    } else {
      locale = null;
    }
    dialog.setVisible(false);
  }

  /**
   * closes the dialog
   */
  public void closeDialog() {
    locale = null;
    dialog.setVisible(false);
    if (debugMode) {
      WtMessageHandler.printToLogFile("AiDialog: closeDialog: Close AI Dialog");
    }
  }
  
  /**
   * returns an array of the translated names of the languages supported by LT
   */
  private String[] getPossibleLanguages() {
    List<String> languages = new ArrayList<>();
    for (Language lang : Languages.get()) {
      languages.add(lang.getTranslatedName(messages));
      if("English".equals(lang.getName())) {
        startLang = new String(lang.getTranslatedName(messages));
      }
    }
    languages.sort(null);
    return languages.toArray(new String[languages.size()]);
  }

  /**
   * returns the locale from a translated language name 
   */
  private Locale getLocaleFromLanguageName(String translatedName) {
    for (Language lang : Languages.get()) {
      if (translatedName.equals(lang.getTranslatedName(messages))) {
        return (WtLinguisticServices.getLocale(lang));
      }
    }
    return null;
  }

  public class TranslationOptions {
    public Locale locale;
    public float temperature;
    
    TranslationOptions(Locale locale, float temperature) {
      this.locale = locale;
      this.temperature = temperature;
    }
  }
  

}
