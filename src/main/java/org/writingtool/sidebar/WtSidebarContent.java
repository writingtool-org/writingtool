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
package org.writingtool.sidebar;

import java.awt.Dimension;
import java.awt.SystemColor;
import java.awt.Toolkit;
import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;
import java.util.SortedMap;
import java.util.TreeMap;

import com.sun.star.accessibility.XAccessible;
import com.sun.star.awt.ActionEvent;
import com.sun.star.awt.ImagePosition;
import com.sun.star.awt.PosSize;
import com.sun.star.awt.Rectangle;
import com.sun.star.awt.WindowEvent;
import com.sun.star.awt.XActionListener;
import com.sun.star.awt.XButton;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XControlModel;
import com.sun.star.awt.XFixedText;
import com.sun.star.awt.XTextComponent;
import com.sun.star.awt.XTextListener;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowListener;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.beans.XMultiPropertySet;
import com.sun.star.lang.DisposedException;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.Locale;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lib.uno.helper.ComponentBase;
import com.sun.star.ui.LayoutSize;
import com.sun.star.ui.XSidebarPanel;
import com.sun.star.ui.XToolPanel;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import org.writingtool.WritingTool;
import org.writingtool.WtDocumentCache;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.aisupport.WtAiParagraphChanging;
import org.writingtool.aisupport.WtAiRemote;
import org.writingtool.WtDocumentsHandler;
import org.writingtool.WtLanguageTool;
import org.writingtool.WtProofreadingError;
import org.writingtool.WtSingleDocument;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtViewCursorTools;

/**
 * Create the window for the WT sidebar panel
 * 
 * @since 1.3
 * @author Fred Kruse
 */
public class WtSidebarContent extends ComponentBase implements XToolPanel, XSidebarPanel {

  public static final String CSS_AWT_UNO_CONTROL_BUTTON = "com.sun.star.awt.UnoControlButton";
  public static final String CSS_AWT_UNO_CONTROL_FIXED_TEXT = "com.sun.star.awt.UnoControlFixedText";
  public static final String CSS_AWT_UNO_CONTROL_EDIT = "com.sun.star.awt.UnoControlEdit";
  public static final String CSS_AWT_UNO_CONTROL_CONTAINER = "com.sun.star.awt.UnoControlContainer";

  
  private static final ResourceBundle messages = WtOfficeTools.getMessageBundle();
  
  private final static String OPENING_FORMAT_SIGN = "[";
  private final static String CLOSING_FORMAT_SIGN = "]";
//  private final static String LINE_BREAK = "\n";
  
//  private final static int LINE_MAX_CHAR = 46;
  
  private final static int MINIMAL_WIDTH = 300;

  private final static int BUTTON_MARGIN_TOP = 5;
  private final static int BUTTON_MARGIN_BOTTOM = 5;
  private final static int BUTTON_MARGIN_LEFT = 5;
  private final static int BUTTON_MARGIN_RIGHT = 5;
  private final static int BUTTON_MARGIN_BETWEEN = 5;

  private final static int BUTTON_WIDTH = 32;
  private final static int BUTTON_CONTAINER_WIDTH = MINIMAL_WIDTH - BUTTON_MARGIN_LEFT - BUTTON_MARGIN_RIGHT;
  private final static int BUTTON_CONTAINER_HEIGHT = BUTTON_WIDTH + BUTTON_MARGIN_TOP + BUTTON_MARGIN_BOTTOM;

  private final static int LABEL_TOP = 2 * BUTTON_MARGIN_TOP + 2 * (BUTTON_CONTAINER_HEIGHT + BUTTON_MARGIN_BETWEEN) + BUTTON_MARGIN_BOTTOM;
  private final static int LABEL_LEFT = BUTTON_MARGIN_LEFT;
  private final static int LABEL_WIDTH = MINIMAL_WIDTH - BUTTON_MARGIN_LEFT - BUTTON_MARGIN_RIGHT;;
  private final static int LABEL_HEIGHT = 20;

  private final static int CONTAINER_MARGIN_TOP = 5;
  private final static int CONTAINER_MARGIN_BOTTOM = 5;
  private final static int CONTAINER_MARGIN_LEFT = 5;
  private final static int CONTAINER_MARGIN_RIGHT = 5;
  private final static int CONTAINER_MARGIN_BETWEEN = 5;
  private final static int CONTAINER_TOP = CONTAINER_MARGIN_TOP + LABEL_TOP + LABEL_HEIGHT;
  private final static int MIN_CONTAINER_HEIGHT = 40;

  private final static int OVERRIDE_BUTTON_WIDTH = 8 * messages.getString("loAiDialogOverrideButton").length();

  private int containerHeight = MIN_CONTAINER_HEIGHT;
  private int OverrideButtonTop = CONTAINER_TOP + 2* MIN_CONTAINER_HEIGHT + 
                                                 3 * CONTAINER_MARGIN_BETWEEN + LABEL_HEIGHT + BUTTON_CONTAINER_HEIGHT;

  private XComponentContext xContext;           //  the component context
  private WtDocumentsHandler documents;
  
//  private XWindow parentWindow;                 //  the parent window
  private XWindow contentWindow;                //  the window of the control container
  private XWindow buttonContainer1Window;       //  the window of the first row button container
  private XWindow buttonContainer2Window;       //  the window of the second row button container
  private XWindow buttonContainerAiWindow;      //  the window of the button container for AI action
  private XWindow paragraphLabelWindow;         //  the window of the paragraph label
  private XWindow paragraphBoxWindow;           //  the window of the paragraph box
  private XWindow aiLabelWindow;                //  the window of the AI label
  private XWindow aiResultBoxWindow = null;     //  the window of the AI result box
  private XWindow overrideButtonWindow;         //  the window of the button to override paragraph with result of AI
  private XWindow buttonAutoOnWindow;           //  the window of button for auto check on
  private XWindow buttonAutoOffWindow;          //  the window of button for auto check off
  private XWindow buttonStatAnWindow;           //  the window of button for statistical Analysis
  private XWindow buttonAiGeneralWindow;        //  the window of button for AI general
  private XWindow buttonAiTranslateTextWindow;  //  the window of button for AI text to speech
  private XWindow buttonAiTextToSpeechWindow;   //  the window of button for AI text to speech
  private XMultiComponentFactory xMCF;          //  The component factory
  private XControlContainer controlContainer;   //  The container of the controls
  private XTextComponent paragraphBox;          //  Box to show paragraph text
  private XTextComponent aiResultBox;           //  Box to show AI result text
  
  private String paragraphText;                 //  Text of the current paragraph
  private String aiResultText;                  //  result Text of AI operation
  private TextParagraph tPara;                  //  Number and type of the current paragraph
  
  boolean isAiSupport = false;
  
  public WtSidebarContent(XComponentContext context, XWindow parentWindow) {
    xContext = context;
//    this.parentWindow = parentWindow;
    try {
      XWindowListener windowAdapter = new XWindowListener() {
        @Override
        public void windowResized(WindowEvent e) {
          resizeContainer();
        }
        @Override
        public void disposing(EventObject e) { }
        @Override
        public void windowHidden(EventObject e) { }
        @Override
        public void windowMoved(WindowEvent e) { }
        @Override
        public void windowShown(EventObject e) { }
      };
      
      parentWindow.addWindowListener(windowAdapter);
  
      xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class, context.getServiceManager());
      XWindowPeer parentWindowPeer = UnoRuntime.queryInterface(XWindowPeer.class, parentWindow);
  
      if (parentWindowPeer == null) {
          return;
      }
  
      XToolkit parentToolkit = parentWindowPeer.getToolkit();
      SortedMap<String, Object> props = new TreeMap<>();
      props.put("BackgroundColor", SystemColor.menu.getRGB() & ~0xFF000000);

      //  Get basic control container
      controlContainer = UnoRuntime.queryInterface(XControlContainer.class, 
          createControlContainer(xMCF, context, new Rectangle(0, 0, 0, 0), props));
      XControl xContainer = UnoRuntime.queryInterface(XControl.class, controlContainer);
      xContainer.createPeer(parentToolkit, parentWindowPeer);
      contentWindow = UnoRuntime.queryInterface(XWindow.class, xContainer);
      
      // Add first row button container
      XControl xButtonContainer = createControlContainer(xMCF, context, 
          new Rectangle(CONTAINER_MARGIN_LEFT, CONTAINER_MARGIN_TOP, BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT), props);
      XControlContainer buttonContainer = UnoRuntime.queryInterface(XControlContainer.class, xButtonContainer);

      controlContainer.addControl("buttonContainer1", xButtonContainer);
      buttonContainer1Window = UnoRuntime.queryInterface(XWindow.class, xButtonContainer);
      buttonContainer1Window.setPosSize(CONTAINER_MARGIN_LEFT, CONTAINER_MARGIN_TOP, BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT, PosSize.POSSIZE);
      
      // Add buttons to container
      int num = 0;
      addButtonToContainer(num, "nextError", "WTNextSmall.png", "loContextMenuNextError", buttonContainer);
      num++;
      addButtonToContainer(num, "checkDialog", "WTCheckSmall.png", "loContextMenuGrammarCheck", buttonContainer);
      num++;
      addButtonToContainer(num, "checkAgainDialog", "WTCheckAgainSmall.png", "loContextMenuGrammarCheckAgain", buttonContainer);
      num++;
      addButtonToContainer(num, "refreshCheck", "WTRefreshSmall.png", "loContextMenuRefreshCheck", buttonContainer);
      num++;
      addButtonToContainer(num, "configure", "WTOptionsSmall.png", "loContextMenuOptions", buttonContainer);
      num++;
      addButtonToContainer(num, "about", "WTAboutSmall.png", "loContextMenuAbout", buttonContainer);
      num++;
      buttonAutoOnWindow = addButtonToContainer(num, "backgroundCheckOn", "WTBackCheckOnSmall.png", "loMenuEnableBackgroundCheck", buttonContainer);
      buttonAutoOffWindow = addButtonToContainer(num, "backgroundCheckOff", "WTBackCheckOffSmall.png", "loMenuDisableBackgroundCheck", buttonContainer);
      documents = WritingTool.getDocumentsHandler();
      buttonAutoOnWindow.setVisible(documents.isBackgroundCheckOff());
      buttonAutoOffWindow.setVisible(!documents.isBackgroundCheckOff());
      
      isAiSupport = documents.getConfiguration().useAiSupport();

      // Add second row button container
      xButtonContainer = createControlContainer(xMCF, context, 
          new Rectangle(CONTAINER_MARGIN_LEFT, CONTAINER_MARGIN_TOP + BUTTON_CONTAINER_HEIGHT + CONTAINER_MARGIN_BETWEEN, 
              BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT), props);
      buttonContainer = UnoRuntime.queryInterface(XControlContainer.class,xButtonContainer);
      controlContainer.addControl("buttonContainer2", xButtonContainer);
      buttonContainer2Window = UnoRuntime.queryInterface(XWindow.class, xButtonContainer);
      buttonContainer2Window.setPosSize(CONTAINER_MARGIN_LEFT, CONTAINER_MARGIN_TOP + BUTTON_CONTAINER_HEIGHT + CONTAINER_MARGIN_BETWEEN, 
          BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT, PosSize.POSSIZE);
      
      // Add buttons to container
      num = 0;
      buttonStatAnWindow = addButtonToContainer(num, "statisticalAnalyses", "WTStatAnSmall.png", "loStatisticalAnalysis", buttonContainer);
      buttonStatAnWindow.setEnable(hasStatAn());
      num++;
      buttonAiGeneralWindow = addButtonToContainer(num, "aiGeneralCommand", "WTAiGeneralSmall.png", "loMenuAiGeneralCommand", buttonContainer);
      buttonAiGeneralWindow.setEnable(isAiSupport);
      num++;
      buttonAiTranslateTextWindow = addButtonToContainer(num, "aiTranslateText", "WTAiTranslateTextSmall.png", "loMenuAiTranslateCommand", buttonContainer);
      buttonAiTranslateTextWindow.setEnable(isAiSupport);
      num++;
      buttonAiTextToSpeechWindow = addButtonToContainer(num, "aiTextToSpeech", "WTAiTextToSpeechSmall.png", "loMenuAiTextToSpeechCommand", buttonContainer);
      buttonAiTextToSpeechWindow.setEnable(documents.getConfiguration().useAiTtsSupport());
      
      // Add Label
      props.put("FontStyleName", "Bold");
      XControl xParagraphLabel = createLabel(xMCF, context, messages.getString("loSidebarParagraphLabel") + ":", 
          new Rectangle(LABEL_LEFT, LABEL_TOP, LABEL_WIDTH, LABEL_HEIGHT), props); 
      controlContainer.addControl("paragraphBoxLabel", xParagraphLabel);
      paragraphLabelWindow = UnoRuntime.queryInterface(XWindow.class, xParagraphLabel);
      
      // Add text field
      props = new TreeMap<>();
      props.put("TextColor", SystemColor.textInactiveText.getRGB() & ~0xFF000000);
      props.put("BackgroundColor", SystemColor.control.getRGB() & ~0xFF000000);
      props.put("Autocomplete", false);
      props.put("HideInactiveSelection", true);
      props.put("MultiLine", true);
      props.put("VScroll", true);
      props.put("Border", (short) 1);
      Rectangle containerSize = contentWindow.getPosSize();
      int paraBoxX = CONTAINER_MARGIN_LEFT;
      int paraBoxY = CONTAINER_TOP;
      int paraBoxWidth = containerSize.Width - CONTAINER_MARGIN_LEFT - CONTAINER_MARGIN_RIGHT;
      int paraBoxHeight = isAiSupport ? containerHeight : containerSize.Height - CONTAINER_MARGIN_TOP - CONTAINER_MARGIN_BOTTOM;
      
      XControl xParagraphBox = createTextfield(xMCF, context, "", 
          new Rectangle(paraBoxX, paraBoxY, paraBoxWidth, paraBoxHeight), props, null);
      paragraphBox = UnoRuntime.queryInterface(XTextComponent.class, xParagraphBox);
             
      controlContainer.addControl("paragraphBox", xParagraphBox);
  
      paragraphBoxWindow = UnoRuntime.queryInterface(XWindow.class, xParagraphBox);
      paragraphBoxWindow.setPosSize(paraBoxX, paraBoxY, paraBoxWidth, paraBoxHeight, PosSize.POSSIZE);
      
      // Add AI Button container
      props = new TreeMap<>();
      props.put("BackgroundColor", SystemColor.menu.getRGB() & ~0xFF000000);
      xButtonContainer = createControlContainer(xMCF, context, 
          new Rectangle(CONTAINER_MARGIN_LEFT, CONTAINER_TOP + containerHeight + CONTAINER_MARGIN_BETWEEN, 
              BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT), props);
      buttonContainer = UnoRuntime.queryInterface(XControlContainer.class, xButtonContainer);
      controlContainer.addControl("buttonContainerAi", xButtonContainer);
      buttonContainerAiWindow = UnoRuntime.queryInterface(XWindow.class, xButtonContainer);
      buttonContainerAiWindow.setPosSize(CONTAINER_MARGIN_LEFT, CONTAINER_TOP + containerHeight + CONTAINER_MARGIN_BETWEEN, 
          BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT, PosSize.POSSIZE);

      // Add AI buttons
      num = 0;
      addButtonToContainer(num, "aiBetterStyle", "WTAiBetterStyleSmall.png", "loMenuAiStyleCommand", buttonContainer, true);
      num++;
      addButtonToContainer(num, "aiReformulateText", "WTAiReformulateSmall.png", "loMenuAiReformulateCommand", buttonContainer, true);
      num++;
      addButtonToContainer(num, "aiAdvanceText", "WTAiExpandSmall.png", "loMenuAiExpandCommand", buttonContainer, true);

      // Add AI Label
      props.put("FontStyleName", "Bold");
      int labelTop = CONTAINER_TOP + containerHeight + 2 * CONTAINER_MARGIN_BETWEEN + BUTTON_CONTAINER_HEIGHT;
      XControl xAiLabel = createLabel(xMCF, context, messages.getString("loAiDialogResultLabel") + ":", 
          new Rectangle(LABEL_LEFT, labelTop, LABEL_WIDTH, LABEL_HEIGHT), props); 
      controlContainer.addControl("aiLabel", xAiLabel);
      aiLabelWindow = UnoRuntime.queryInterface(XWindow.class, xAiLabel);
      
      // Add text field
      props = new TreeMap<>();
      props.put("TextColor", SystemColor.textInactiveText.getRGB() & ~0xFF000000);
      props.put("BackgroundColor", SystemColor.control.getRGB() & ~0xFF000000);
      props.put("Autocomplete", false);
      props.put("HideInactiveSelection", true);
      props.put("MultiLine", true);
      props.put("VScroll", true);
      props.put("Border", (short) 1);
      int aiBoxX = CONTAINER_MARGIN_LEFT;
      int aiBoxY = CONTAINER_TOP + containerHeight + 2 * CONTAINER_MARGIN_BETWEEN + LABEL_HEIGHT + BUTTON_CONTAINER_HEIGHT;
      int aiBoxWidth = containerSize.Width - CONTAINER_MARGIN_LEFT - CONTAINER_MARGIN_RIGHT;
      int aiBoxHeight = containerHeight;
      XControl xAiBox = createTextfield(xMCF, context, "", 
          new Rectangle(aiBoxX, aiBoxY, aiBoxWidth, aiBoxHeight), props, null);
      aiResultBox = UnoRuntime.queryInterface(XTextComponent.class, xAiBox);
      controlContainer.addControl("aiBox", xAiBox);
      aiResultBoxWindow = UnoRuntime.queryInterface(XWindow.class, xAiBox);
      aiResultBoxWindow.setPosSize(aiBoxX, aiBoxY, aiBoxWidth, aiBoxHeight, PosSize.POSSIZE);
      
      // Add override button
      XControl overrideButton = createButton(xMCF, context, messages.getString("loAiDialogOverrideButton"), null, 
          new Rectangle(CONTAINER_MARGIN_LEFT, OverrideButtonTop, OVERRIDE_BUTTON_WIDTH, BUTTON_WIDTH), null);
      XButton xOverrideButton = UnoRuntime.queryInterface(XButton.class, overrideButton);
      XActionListener xoverrideButtonAction = new XActionListener() {
        @Override
        public void actionPerformed(ActionEvent event) {
          try {
          WtAiParagraphChanging.insertText(aiResultText, documents.getCurrentDocument().getXComponent(), true);
          } catch (Throwable e1) {
            WtMessageHandler.printException(e1);
          }
        }
        @Override
        public void disposing(EventObject arg0) { }
      };
      xOverrideButton.addActionListener(xoverrideButtonAction);
      controlContainer.addControl("overrideButton", overrideButton);
      overrideButtonWindow = UnoRuntime.queryInterface(XWindow.class, overrideButton);
      overrideButtonWindow.setPosSize(CONTAINER_MARGIN_LEFT, OverrideButtonTop, OVERRIDE_BUTTON_WIDTH, BUTTON_WIDTH, PosSize.POSSIZE);;
      buttonContainerAiWindow.setVisible(isAiSupport);
      aiLabelWindow.setVisible(isAiSupport);
      aiResultBoxWindow.setVisible(isAiSupport);
      overrideButtonWindow.setVisible(isAiSupport);
      documents.setSidebarContent(this);
      setTextToBox();
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    }
      
  }
  
  /**
   * get the WtSidebar
   */
  public WtSidebarContent getWTSidebar() {
    return this;
  }
  
  /**
   * set AI support
   */
  public void setAiSupport(boolean isAiSupport) {
    this.isAiSupport = isAiSupport;
    resizeContainer();
  }
  
  /**
   * Set the paragraph Text under the Cursor to the sidebar text box
   */
  public void setCursorTextToBox() {
    setCursorTextToBox(null, false);
  }

  public void setCursorTextToBox(boolean force) {
    setCursorTextToBox(null, force);
  }

  public void setCursorTextToBox(XComponent xComponent) {
    setCursorTextToBox(xComponent, false);
  }

  public void setCursorTextToBox(XComponent xComponent, boolean force) {
    try {
      if (xComponent == null) {
        xComponent = WtOfficeTools.getCurrentComponent(xContext);
      }
      WtViewCursorTools vCursor = new WtViewCursorTools(xComponent);
      String pText = vCursor.getViewCursorParagraphText();
      if (pText == null || (!force && pText.equals(paragraphText))) {
        return;
      }
      boolean isSameText = pText.equals(paragraphText);
      paragraphText = pText;
      tPara = vCursor.getViewCursorParagraph();
      WtProofreadingError[] errors = getErrorsOfParagraph(tPara);
      String formatedText = formatText(paragraphText, errors);
      if (formatedText == null) {
        formatedText = "";
      }
      paragraphBox.setText(formatedText);
      if (isAiSupport && aiResultBox != null && !isSameText) {
        aiResultBox.setText("");
      }
    } catch (Throwable e1) {
      WtMessageHandler.showError(e1);
    }
  }
  
/**
 * set a AI result to the AI result box
 */
  public void setTextToAiResultBox(String paraText, String resultText) {
    if (paraText.equals(paragraphText)) {
      aiResultText = resultText;
      aiResultBox.setText(aiResultText);
    }
  }
  
  /**
   * Create a button with label and action listener.
   */
  private static XControl createButton(XMultiComponentFactory xMCF, XComponentContext context, String label,
      XActionListener listener, Rectangle size, SortedMap<String, Object> props) {
    XControl buttonCtrl = createControl(xMCF, context, CSS_AWT_UNO_CONTROL_BUTTON, props, size);
    XButton button = UnoRuntime.queryInterface(XButton.class, buttonCtrl);
    button.setLabel(label);
    if (listener != null) {
      button.addActionListener(listener);
    }
    return buttonCtrl;
  }

  /**
   * Create a label.
   */
  private static XControl createLabel(XMultiComponentFactory xMCF, XComponentContext context, String text,
      Rectangle size, SortedMap<String, Object> props) {
    if (props == null) {
      props = new TreeMap<>();
    }
    XControl buttonCtrl = createControl(xMCF, context, CSS_AWT_UNO_CONTROL_FIXED_TEXT, props, size);
    XFixedText txt = UnoRuntime.queryInterface(XFixedText.class, buttonCtrl);
    txt.setText(text);
    return buttonCtrl;
  }

  /**
   * Create a text field.
   */
  private static XControl createTextfield(XMultiComponentFactory xMCF, XComponentContext context, String text,
      Rectangle size, SortedMap<String, Object> props, XTextListener textListener) {
    if (props == null) {
      props = new TreeMap<>();
    }
//    props.put("VScroll", false);
    XControl buttonCtrl = createControl(xMCF, context, CSS_AWT_UNO_CONTROL_EDIT, props, size);
    XTextComponent txt = UnoRuntime.queryInterface(XTextComponent.class, buttonCtrl);
    if (textListener != null)
      txt.addTextListener(textListener);
    txt.setText(text);
    return buttonCtrl;
  }

  /**
   * Create a control container.
   */
  private static XControl createControlContainer(XMultiComponentFactory xMCF, XComponentContext context, Rectangle size,
      SortedMap<String, Object> props) {
    return createControl(xMCF, context, CSS_AWT_UNO_CONTROL_CONTAINER, props, size);
  }

  /**
   * Creates any control.
   */
  private static XControl createControl(XMultiComponentFactory xMCF, XComponentContext xContext, String type,
      SortedMap<String, Object> props, Rectangle rectangle) {
    try
    {
      XControl control = UnoRuntime.queryInterface(XControl.class, xMCF.createInstanceWithContext(type, xContext));
      XControlModel controlModel = UnoRuntime.queryInterface(XControlModel.class,
          xMCF.createInstanceWithContext(type + "Model", xContext));
      control.setModel(controlModel);
      XMultiPropertySet properties = UnoRuntime.queryInterface(XMultiPropertySet.class, control.getModel());
      if (props != null && props.size() > 0)
      {
        properties.setPropertyValues(props.keySet().toArray(new String[props.size()]),
            props.values().toArray(new Object[props.size()]));
      }
      XWindow controlWindow = UnoRuntime.queryInterface(XWindow.class, control);
      controlWindow.setPosSize(rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height, PosSize.POSSIZE);
      return control;
    } catch (com.sun.star.uno.Exception ex) {
      WtMessageHandler.printException(ex);
      return null;
    }
  }


  
  private XWindow addButtonToContainer(int num, String cmd, String imageFile, String helpText, XControlContainer buttonContainer) {
    return addButtonToContainer(num, cmd, imageFile, helpText, buttonContainer, false);
  }
  
  private XWindow addButtonToContainer(int num, String cmd, String imageFile, String helpText, 
      XControlContainer buttonContainer, boolean isAiCmd) {
    SortedMap<String, Object> bProps = new TreeMap<>();
    bProps.put("ImageURL", "vnd.sun.star.extension://org.openoffice.writingtool.oxt/images/" + imageFile);
    bProps.put("HelpText", messages.getString(helpText));
    bProps.put("ImagePosition", ImagePosition.Centered);
    XControl button = createButton(xMCF, xContext, "", null, new Rectangle(0, 0, 16, 16), bProps);
    XButton xButton = UnoRuntime.queryInterface(XButton.class, button);
    XActionListener xButtonOptionsAction = new XActionListener() {
      @Override
      public void actionPerformed(ActionEvent event) {
        try {
          if (isAiCmd) {
            runAICommand(cmd);
          } else {
            runGeneralDispatchCmd(cmd);
          }
        } catch (Throwable e1) {
          WtMessageHandler.showError(e1);
        }
      }
      @Override
      public void disposing(EventObject arg0) { }
    };
    xButton.addActionListener(xButtonOptionsAction);
    buttonContainer.addControl(cmd, button);
    XWindow xWindow = UnoRuntime.queryInterface(XWindow.class, button);
    xWindow.setPosSize(BUTTON_MARGIN_LEFT + num * (BUTTON_WIDTH + BUTTON_MARGIN_BETWEEN), 
        BUTTON_MARGIN_TOP, BUTTON_WIDTH, BUTTON_WIDTH, PosSize.POSSIZE);
    return xWindow;
  }
  
  private void runAICommand(String cmd) throws Throwable {
    WtAiRemote aiRemote = new WtAiRemote(documents, documents.getConfiguration());
    WtDocumentCache docCache = documents.getCurrentDocument().getDocumentCache();
    Locale locale = docCache.getTextParagraphLocale(tPara);
    String instruction = null;
    boolean onlyPara = false;
    float temp = 0.7f;
    if (cmd.equals("aiBetterStyle")) {
      instruction = WtAiRemote.getInstruction(WtAiRemote.STYLE_INSTRUCTION, locale);
      onlyPara = true;
      temp = 0.0f;
    } else if (cmd.equals("aiReformulateText")) {
      instruction = WtAiRemote.getInstruction(WtAiRemote.REFORMULATE_INSTRUCTION, locale);
      onlyPara = true;
    } else if (cmd.equals("aiAdvanceText")) {
      instruction = WtAiRemote.getInstruction(WtAiRemote.EXPAND_INSTRUCTION, locale);
    } else {
      WtMessageHandler.showMessage("Unknown command: " + cmd);
    }
    aiResultText = aiRemote.runInstruction(instruction, paragraphText, temp, 1, locale, onlyPara);
    aiResultBox.setText(aiResultText);
  }
  
  private void runGeneralDispatchCmd(String cmd) throws Throwable {
    WtSingleDocument document = documents.getCurrentDocument();
    if (document != null) {
      document.setMenuDocId();
    }
    WtOfficeTools.dispatchCmd("service:org.writingtool.WritingTool?" + cmd, xContext);
    setTextToBox();
  }

  private boolean hasStatAn() {
    return WtOfficeTools.hasStatisticalStyleRules(xContext);
  }
  
  private void resizeContainer() {
    Rectangle rect = contentWindow.getPosSize();
    containerHeight = rect.Height - CONTAINER_TOP - CONTAINER_MARGIN_BOTTOM;
//    WtMessageHandler.printToLogFile("rect.Height: " + rect.Height + ", containerHeight: " + containerHeight);
    if (isAiSupport) {
      containerHeight = (containerHeight - 2 * BUTTON_CONTAINER_HEIGHT - 4 * CONTAINER_MARGIN_BETWEEN - LABEL_HEIGHT) / 2;
    }
    if (containerHeight < 0) {
      containerHeight = MIN_CONTAINER_HEIGHT;
    }
//    contentWindow.setPosSize(rect.X, rect.Y, rect.Width, rect.Height, PosSize.POSSIZE);
    buttonContainer1Window.setPosSize(CONTAINER_MARGIN_LEFT, CONTAINER_MARGIN_TOP, 
        BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT, PosSize.POSSIZE);
    buttonContainer2Window.setPosSize(CONTAINER_MARGIN_LEFT, CONTAINER_MARGIN_TOP + BUTTON_CONTAINER_HEIGHT + CONTAINER_MARGIN_BETWEEN, 
        BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT, PosSize.POSSIZE);
    paragraphLabelWindow.setPosSize(LABEL_LEFT, LABEL_TOP, LABEL_WIDTH, LABEL_HEIGHT, PosSize.POSSIZE);
    int paraBoxX = CONTAINER_MARGIN_LEFT;
    int paraBoxY = CONTAINER_TOP;
    int paraBoxWidth = rect.Width - CONTAINER_MARGIN_LEFT - CONTAINER_MARGIN_RIGHT;
    int paraBoxHeight = containerHeight;
    paragraphBoxWindow.setPosSize(paraBoxX, paraBoxY, paraBoxWidth, paraBoxHeight, PosSize.POSSIZE);
    buttonAutoOnWindow.setVisible(documents.isBackgroundCheckOff());
    buttonAutoOffWindow.setVisible(!documents.isBackgroundCheckOff());
    buttonStatAnWindow.setEnable(hasStatAn());
    buttonAiGeneralWindow.setEnable(isAiSupport);
    buttonAiTranslateTextWindow.setEnable(isAiSupport);
    buttonAiTextToSpeechWindow.setEnable(documents.getConfiguration().useAiTtsSupport());
    if (isAiSupport) {
      OverrideButtonTop = CONTAINER_TOP + 2* containerHeight + 
          3 * CONTAINER_MARGIN_BETWEEN + LABEL_HEIGHT + BUTTON_CONTAINER_HEIGHT;
      buttonContainerAiWindow.setPosSize(CONTAINER_MARGIN_LEFT, CONTAINER_TOP + containerHeight + CONTAINER_MARGIN_BETWEEN, 
          BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT, PosSize.POSSIZE);
      aiLabelWindow.setPosSize(LABEL_LEFT, CONTAINER_TOP + containerHeight + 2 * CONTAINER_MARGIN_BETWEEN + BUTTON_CONTAINER_HEIGHT, 
          LABEL_WIDTH, LABEL_HEIGHT, PosSize.POSSIZE);
      int aiBoxX = CONTAINER_MARGIN_LEFT;
      int aiBoxY = CONTAINER_TOP + containerHeight + 2 * CONTAINER_MARGIN_BETWEEN + LABEL_HEIGHT + BUTTON_CONTAINER_HEIGHT;
      int aiBoxWidth = rect.Width - CONTAINER_MARGIN_LEFT - CONTAINER_MARGIN_RIGHT;
      int aiBoxHeight = containerHeight;
      aiResultBoxWindow.setPosSize(aiBoxX, aiBoxY, aiBoxWidth, aiBoxHeight, PosSize.POSSIZE);
      overrideButtonWindow.setPosSize(CONTAINER_MARGIN_LEFT, OverrideButtonTop, OVERRIDE_BUTTON_WIDTH, BUTTON_WIDTH, PosSize.POSSIZE);;
    }
    buttonContainerAiWindow.setVisible(isAiSupport);
    aiLabelWindow.setVisible(isAiSupport);
    aiResultBoxWindow.setVisible(isAiSupport);
    overrideButtonWindow.setVisible(isAiSupport);
  }
  
  private WtProofreadingError[] getErrorsOfParagraph(TextParagraph tPara) {
    WtLanguageTool lt = documents.getLanguageTool();
    WtSingleDocument document = documents.getCurrentDocument();
    if (document == null) {
      return new WtProofreadingError[0];
    }
    WtDocumentCache docCache = document.getDocumentCache();
    if (docCache == null) {
      return new WtProofreadingError[0];
    }
    int nFPara = docCache.getFlatParagraphNumber(tPara);
    List<WtProofreadingError[]> errors = new ArrayList<>();
    for (int cacheNum = 0; cacheNum < document.getParagraphsCache().size(); cacheNum++) {
      if (!docCache.isAutomaticGenerated(nFPara, true) && ((cacheNum == WtOfficeTools.CACHE_SINGLE_PARAGRAPH 
              || (lt.isSortedRuleForIndex(cacheNum) && !document.getDocumentCache().isSingleParagraph(nFPara)))
          || (cacheNum == WtOfficeTools.CACHE_AI && documents.useAi()))) {
        WtProofreadingError[] pErrors = document.getParagraphsCache().get(cacheNum).getSafeMatches(nFPara);
        //  Note: unsafe matches are needed to prevent the thread to get into a read lock
        errors.add(pErrors);
      } else {
        errors.add(new WtProofreadingError[0]);
      }
    }
    return document.mergeErrors(errors, nFPara);
  }
  
  private String formatText(String orgText, WtProofreadingError[] errors) {
    if (errors == null || errors.length == 0) {
      return orgText;
    }
    String formatedText = "";
    int lastChar = 0;
    for (WtProofreadingError error : errors) {
      if (lastChar <= error.nErrorStart && error.nErrorStart + error.nErrorLength <= orgText.length()) {
        if (error.nErrorStart > 0 &&  error.nErrorStart <= orgText.length()) {
          formatedText += orgText.substring(lastChar, error.nErrorStart);
        }
        lastChar = error.nErrorStart + error.nErrorLength;
        formatedText += OPENING_FORMAT_SIGN + orgText.substring(error.nErrorStart, lastChar) + CLOSING_FORMAT_SIGN;
      }
    }
    if (lastChar < orgText.length()) {
      formatedText += orgText.substring(lastChar);
    }
    return formatedText;
  }
  
  private void setTextToBox() {
    try {
      resizeContainer();
      setCursorTextToBox();
    } catch (Throwable e1) {
      WtMessageHandler.showError(e1);
    }
  }
/*  
  private String addLineBreaks(String inpText) {
    if (inpText.length() <= LINE_MAX_CHAR) {
      return inpText;
    }
    String outpText = "";
    int lastChar = 0;
    while (lastChar < inpText.length()) {
      if (lastChar + LINE_MAX_CHAR >= inpText.length()) {
        outpText += inpText.substring(lastChar);
        lastChar = inpText.length();
      } else {
        int i;
        for (i = lastChar + LINE_MAX_CHAR; i > lastChar && !Character.isWhitespace(inpText.charAt(i)) && inpText.charAt(i) != '-'; i--);
//        WtMessageHandler.printToLogFile("lastChar = " + ", i = " + i);
        if (i == 0) {
          outpText += inpText.substring(lastChar, lastChar + LINE_MAX_CHAR) + LINE_BREAK;
          lastChar += LINE_MAX_CHAR;
        } else {
          outpText += inpText.substring(lastChar, i + 1) + LINE_BREAK;
          lastChar = i + 1;
        }
      }
    }
    return outpText;
  }
*/
  @Override
  public XAccessible createAccessible(XAccessible a) {
    return UnoRuntime.queryInterface(XAccessible.class, getWindow());
  }

  @Override
  public XWindow getWindow() {
    if (controlContainer == null) {
      throw new DisposedException("", this);
    }
    return UnoRuntime.queryInterface(XWindow.class, controlContainer);
  }

  @Override
  public LayoutSize getHeightForWidth(int width) {
//    int height = parentWindow.getPosSize().Height - CONTAINER_MARGIN_TOP - CONTAINER_MARGIN_BOTTOM;
    Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
    int minHeight = CONTAINER_TOP + 2* MIN_CONTAINER_HEIGHT + 
        3 * CONTAINER_MARGIN_BETWEEN + LABEL_HEIGHT + 2 * BUTTON_CONTAINER_HEIGHT + CONTAINER_MARGIN_BOTTOM ;
    int maxHeight = (int) (0.65 * screenSize.height) - 2 * CONTAINER_MARGIN_BOTTOM;
    return new LayoutSize(minHeight, maxHeight, maxHeight);
//    int height = OverrideButtonTop + BUTTON_WIDTH + CONTAINER_MARGIN_BOTTOM;
//    return new LayoutSize(height, height, height);
  }

//  @Override
  public int getMinimalWidth() {
    return MINIMAL_WIDTH;
  }
  
  
}
