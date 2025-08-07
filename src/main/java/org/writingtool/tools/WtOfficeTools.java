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

import static org.languagetool.JLanguageTool.MESSAGE_BUNDLE;

import java.awt.Image;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.MissingResourceException;
import java.util.ResourceBundle;
import java.util.TreeMap;

import javax.imageio.ImageIO;
import javax.swing.ImageIcon;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.SystemUtils;
import org.jetbrains.annotations.Nullable;
import org.languagetool.JLanguageTool;
import org.languagetool.Language;
import org.languagetool.LanguageMaintainedState;
import org.languagetool.Languages;
import org.languagetool.language.Contributor;
import org.languagetool.rules.AbstractStatisticSentenceStyleRule;
import org.languagetool.rules.AbstractStatisticStyleRule;
import org.languagetool.rules.AbstractStyleTooOftenUsedWordRule;
import org.languagetool.rules.ReadabilityRule;
import org.languagetool.rules.Rule;
import org.writingtool.WtDictionary;
import org.writingtool.WtDocumentsHandler;
import org.writingtool.WtProofreadingError;
import org.writingtool.languagedetectors.WtKhmerDetector;
import org.writingtool.languagedetectors.WtTamilDetector;

import com.sun.star.awt.XMenuBar;
import com.sun.star.awt.XPopupMenu;
import com.sun.star.awt.XWindow;
import com.sun.star.beans.Property;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.frame.XController;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XDispatchHelper;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XLayoutManager;
import com.sun.star.frame.XModel;
import com.sun.star.lang.DisposedException;
import com.sun.star.lang.Locale;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.linguistic2.SingleProofreadingError;
import com.sun.star.linguistic2.XProofreadingIterator;
import com.sun.star.linguistic2.XSearchableDictionaryList;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextViewCursor;
import com.sun.star.text.XTextViewCursorSupplier;
import com.sun.star.ui.XUIElement;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
 * Some tools to get information of LibreOffice/OpenOffice document context
 * @since 1.0
 * @author Fred Kruse
 */
public class WtOfficeTools {
  
  public enum DocumentType {
    WRITER,       //  Writer document
    IMPRESS,      //  Impress document
    CALC,         //  Calc document
    UNSUPPORTED   //  unsupported document
  }
  
  public enum RemoteCheck {
    NONE,         //  no remote check
    ALL,          //  spell and grammar check
    ONLY_SPELL,   //  only spell check
    ONLY_GRAMMAR  //  only grammar check
  }
    
  public enum LoErrorType {
    GRAMMAR,      //  grammar error
    SPELL,        //  spell error
    BOTH          //  spell and grammar error
  }
    
  public static final String WT_NAME = "WritingTool";

  public static final String AI_GRAMMAR_CATEGORY = "AI_GRAMMAR_CATEGORY";
  public static final String AI_STYLE_CATEGORY = "AI_STYLE_CATEGORY";
  public static final String AI_GRAMMAR_RULE_ID = "LO_AI_DETECTION_RULE";
  public static final String AI_UNKNOWN_WORD_RULE_ID = "LO_AI_DETECTION_RULE_UNKNOWN_WORD";
  public static final String AI_GRAMMAR_HINT_RULE_ID = "LO_AI_DETECTION_RULE_HINT";
  public static final String AI_GRAMMAR_OTHER_RULE_ID = "LO_AI_DETECTION_RULE_OTHER";

  public static final String WT_SERVER = "https://writingtool.org";
  public static final String WT_SERVER_URL = WT_SERVER + "/index.php";
  public static final String EXTENSION_MAINTAINER = "Fred Kruse";
  public static final String WT_SERVICE_NAME = "org.writingtool.WritingTool";
  public static final String WT_SPELL_SERVICE_NAME = "org.writingtool.WritingToolSpellChecker";
  public static final String WT_DISPLAY_SERVICE_NAME = WT_NAME;
  private final static String RESOURCE_PATH = "/org/writingtool/resource/";
  
  
  public static final int PROOFINFO_UNKNOWN = 0;
  public static final int PROOFINFO_GET_PROOFRESULT = 1;
  public static final int PROOFINFO_MARK_PARAGRAPH = 2;

  public static final int NUMBER_TEXTLEVEL_CACHE = 4;     // Number of caches for matches of text level rules
  public static final int NUMBER_CACHE = 5;               // Number of all caches
  public static final int CACHE_SINGLE_PARAGRAPH = 0;     // Cache for matches of sentences and single paragraph rules
  public static final int CACHE_N_PARAGRAPH = 1;          // Cache for matches of n paragraph rules
  public static final int CACHE_CHAPTER = 2;              // Cache for matches of chapter rules
  public static final int CACHE_TEXT = 3;                 // Cache for matches of whole text rules
  public static final int CACHE_AI = 4;                   // Cache for matches of AI proofs

  public static final String END_OF_PARAGRAPH = "\n\n";   //  Paragraph Separator like in standalone GUI
  public static final int NUMBER_PARAGRAPH_CHARS = END_OF_PARAGRAPH.length();  //  number of end of paragraph characters
  public static final String SINGLE_END_OF_PARAGRAPH = "\n";
  public static final String MANUAL_LINEBREAK = "\r";     //  to distinguish from paragraph separator
  public static final String ZERO_WIDTH_SPACE = "\u200B"; // Used to mark footnotes, functions, etc.
  public static final char ZERO_WIDTH_SPACE_CHAR = '\u200B'; // Used to mark footnotes, functions, etc.
  public static final String SOFT_HYPHEN = "\u00AD";      // Soft Hyphen (has to be removed for grammar check)
  public static final String IGNORE_LANGUAGE = "zxx";     // Used from LT to mark automatic generated text like indexes
  public static final String LOG_LINE_BREAK = System.lineSeparator();  //  LineBreak in Log-File (MS-Windows compatible)
  public static final int MAX_SUGGESTIONS = 25;           // Number of suggestions maximal shown in LO/OO
  public static final String MULTILINGUAL_LABEL = "99-";  // Label added in front of variant to indicate a multilingual paragraph (returned is the main language)
  public static final int CHECK_MULTIPLIKATOR = 40;       //  Number of minimum checks for a first check run
  public static final int CHECK_SHAPES_TIME = 1000;       //  time interval to run check for changes in text inside of shapes
  public static final int SPELL_CHECK_MIN_HEAP = 850;     //  Minimal heap space to run LT spell check
  public static int TIME_TOLERANCE = 100;                 //  Minimal milliseconds to show message in TM debug mode
  
  public static int DEBUG_MODE_SD = 0;            //  Set Debug Mode for SingleDocument
  public static int DEBUG_MODE_SC = 0;            //  Set Debug Mode for SingleCheck
  public static int DEBUG_MODE_CR = 0;            //  Set Debug Mode for CheckRequest
  public static int DEBUG_MODE_AI = 0;            //  Set Debug Mode for AI support
  public static boolean DEBUG_MODE_CD = false;    //  Activate Debug Mode for SpellAndGrammarCheckDialog
  public static boolean DEBUG_MODE_DC = false;    //  Activate Debug Mode for DocumentCache
  public static boolean DEBUG_MODE_FP = false;    //  Activate Debug Mode for FlatParagraphTools
  public static boolean DEBUG_MODE_IO = false;    //  Activate Debug Mode for Cache save to file
  public static boolean DEBUG_MODE_LD = false;    //  Activate Debug Mode for LtDictionary
  public static boolean DEBUG_MODE_LM = false;    //  Activate Debug Mode for LanguageToolMenus
  public static boolean DEBUG_MODE_MD = false;    //  Activate Debug Mode for MultiDocumentsHandler
  public static boolean DEBUG_MODE_RM = false;    //  Activate Debug Mode for RemoteLanguageTool
  public static boolean DEBUG_MODE_SP = false;    //  Activate Debug Mode for LtSpellChecker
  public static boolean DEBUG_MODE_SR = false;    //  Activate Debug Mode for SortedTextRules
  public static boolean DEBUG_MODE_TA = false;    //  Activate Debug Mode for time measurements (only AI)
  public static boolean DEBUG_MODE_TM = false;    //  Activate Debug Mode for time measurements
  public static boolean DEBUG_MODE_TQ = false;    //  Activate Debug Mode for TextLevelCheckQueue
  public static boolean DEVELOP_MODE_ST = false;  //  Activate Development Mode to test sorted text IDs
  public static boolean DEVELOP_MODE = false;     //  Activate Development Mode

  public  static final String CONFIG_FILE = "WritingTool.cfg";
  public  static final String OOO_CONFIG_FILE = "WritingTool-ooo.cfg";
  private static final String LOG_FILE = "WritingTool.log";
  private static final String LOG_FILE_SP = "WritingToolSpell.log";
  public  static final String STATISTICAL_ANALYZES_CONFIG_FILE = "WT_Statistical_Analyzes.cfg";

  private static final String VENDOR_ID = "writingtool.org";
  private static final String APPLICATION_ID = "WritingTool";
  private static final String CACHE_ID = "cache";
  private static String OFFICE_EXTENSION_ID = null;

  private static final String RESOURCES = "org.writingtool";
  
  private static final String MENU_BAR = "private:resource/menubar/menubar";
  private static final String LOG_DELIMITER = ",";
  

  private static final double LT_HEAP_LIMIT_FACTOR = 0.9;
  private static double MAX_HEAP_SPACE = -1;
  private static double LT_HEAP_LIMIT = -1;
  
  private final static int MAX_LO_WAITS = 3000;
  private static int numLoWaits = 0;
  private static Object waitObj = new Object();

/*
  private static final long KEY_RELEASE_TOLERANCE = 500;
  private static long lastKeyRelease = 0;
*/
  
  /**
   * Returns the XDesktop
   * Returns null if it fails
   */
  @Nullable
  public static XDesktop getDesktop(XComponentContext xContext) {
    try {
      if (xContext == null) {
        return null;
      }
      XMultiComponentFactory xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class,
              xContext.getServiceManager());
      if (xMCF == null) {
        return null;
      }
      Object desktop = xMCF.createInstanceWithContext("com.sun.star.frame.Desktop", xContext);
      if (desktop == null) {
        return null;
      }
      return UnoRuntime.queryInterface(XDesktop.class, desktop);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return null;           // Return null as method failed
    }
  }

  /** 
   * Returns the current XFrame 
   * Returns null if it fails
   */
  @Nullable
  public static XFrame getCurrentFrame(XComponentContext xContext) {
    XDesktop xDesktop = getDesktop(xContext);
    if (xDesktop == null) {
      return null;
    }
    return xDesktop.getCurrentFrame();
  }

  /** 
   * Returns the current XFrame 
   * Returns null if it fails
   */
  @Nullable
  public static XWindow getCurrentWindow(XComponentContext xContext) {
    XFrame xFrame = getCurrentFrame(xContext);
    if (xFrame == null) {
      return null;
    }
    return xFrame.getContainerWindow();
  }
  
  /** 
   * Returns the current XComponent 
   * Returns null if it fails
   */
  @Nullable
  public static XComponent getCurrentComponent(XComponentContext xContext) {
    try {
      XDesktop xdesktop = getDesktop(xContext);
      if (xdesktop == null) {
        return null;
      }
      else return xdesktop.getCurrentComponent();
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return null;           // Return null as method failed
    }
  }
    
  /**
   * Returns the current text document (if any) 
   * Returns null if it fails
   */
  @Nullable
  static XTextDocument getCurrentDocument(XComponentContext xContext) {
    try {
      XComponent curComp = getCurrentComponent(xContext);
      if (curComp == null) {
        return null;
      }
      else return UnoRuntime.queryInterface(XTextDocument.class, curComp);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return null;           // Return null as method failed
    }
  }
  
  /**
   * returns the default language of the text document
   */
  public static Language getDefaultLanguage(XComponentContext xContext) {
    Locale locale = getDefaultLocale(xContext);
    if (locale == null) {
      return null;
    }
    return WtDocumentsHandler.getLanguage(locale);
  }

  /**
   * returns the default locale of the text document
   */
  public static Locale getDefaultLocale(XComponentContext xContext) {
    XTextDocument curDoc = getCurrentDocument(xContext);
    if (curDoc == null) {
      return null;
    }
    XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, curDoc);
    if (xPropertySet == null) {
      return null;
    }
    try {
      return (Locale) xPropertySet.getPropertyValue("CharLocale");
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return null;
    }
  }
  
  /**
   * Checks the locale under the cursor. Used for opening the configuration dialog.
   * @return the locale under the visible cursor
   */
  @Nullable
  public static Locale getCursorLocale(XComponentContext xContext) {
    if (xContext == null) {
      return null;
    }
    XComponent xComponent = WtOfficeTools.getCurrentComponent(xContext);
    if (xComponent == null) {
      return null;
    }
    Locale charLocale;
    XPropertySet xCursorProps;
    try {
      //  Test for Impress or Calc document
      if (WtOfficeDrawTools.isImpressDocument(xComponent)) {
        return WtOfficeDrawTools.getDocumentLocale(xComponent);
      } else if (WtOfficeSpreadsheetTools.isSpreadsheetDocument(xComponent)) {
        return WtOfficeSpreadsheetTools.getDocumentLocale(xComponent);
      }
      XModel model = UnoRuntime.queryInterface(XModel.class, xComponent);
      if (model == null) {
        return null;
      }
      XTextViewCursorSupplier xViewCursorSupplier =
          UnoRuntime.queryInterface(XTextViewCursorSupplier.class, model.getCurrentController());
      if (xViewCursorSupplier == null) {
        return null;
      }
      XTextViewCursor xCursor = xViewCursorSupplier.getViewCursor();
      if (xCursor == null) {
        return null;
      }
      if (xCursor.isCollapsed()) { // no text selection
        xCursorProps = UnoRuntime.queryInterface(XPropertySet.class, xCursor);
      } else { // text is selected, need to create another cursor
        // as multiple languages can occur here - we care only
        // about character under the cursor, which might be wrong
        // but it applies only to the checking dialog to be removed
        xCursorProps = UnoRuntime.queryInterface(
            XPropertySet.class,
            xCursor.getText().createTextCursorByRange(xCursor.getStart()));
      }

      // The CharLocale and CharLocaleComplex properties may both be set, so we still cannot know
      // whether the text is e.g. Khmer or Tamil (the only "complex text layout (CTL)" languages we support so far).
      // Thus we check the text itself:
      if (new WtKhmerDetector().isThisLanguage(xCursor.getText().getString())) {
        return new Locale("km", "", "");
      }
      if (new WtTamilDetector().isThisLanguage(xCursor.getText().getString())) {
        return new Locale("ta","","");
      }
      if (xCursorProps == null) {
        return null;
      }
      Object obj = xCursorProps.getPropertyValue("CharLocale");
      if (obj == null) {
        return null;
      }
      charLocale = (Locale) obj;
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      return null;
    }
    return charLocale;
  }

  static void printPropertySet (Object o) {
    XPropertySet propSet = UnoRuntime.queryInterface(XPropertySet.class, o);
    if (propSet == null) {
      WtMessageHandler.printToLogFile("OfficeTools: printPropertySet: XPropertySet == null");
      return;
    }
    XPropertySetInfo propertySetInfo = propSet.getPropertySetInfo();
    WtMessageHandler.printToLogFile("OfficeTools: printPropertySet: PropertySet:");
    for (Property property : propertySetInfo.getProperties()) {
      WtMessageHandler.printToLogFile("Name: " + property.Name + ", Type: " + property.Type.getTypeName());
    }
  }
  
  /**
   * Returns the searchable dictionary list
   * Returns null if it fails
   */
  @Nullable
  public
  static XSearchableDictionaryList getSearchableDictionaryList(XComponentContext xContext) {
    try {
      if (xContext == null) {
        return null;
      }
      XMultiComponentFactory xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class,
              xContext.getServiceManager());
      if (xMCF == null) {
        return null;
      }
      Object dictionaryList = xMCF.createInstanceWithContext("com.sun.star.linguistic2.DictionaryList", xContext);
      if (dictionaryList == null) {
        return null;
      }
      return UnoRuntime.queryInterface(XSearchableDictionaryList.class, dictionaryList);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return null;           // Return null as method failed
    }
  }

  /**
   * Returns the searchable dictionary list
   * Returns null if it fails
   */
  @Nullable
  static XProofreadingIterator getProofreadingIterator(XComponentContext xContext) {
    try {
      if (xContext == null) {
        return null;
      }
      XMultiComponentFactory xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class,
              xContext.getServiceManager());
      if (xMCF == null) {
        return null;
      }
      Object proofreadingIterator = xMCF.createInstanceWithContext("com.sun.star.linguistic2.ProofreadingIterator", xContext);
      if (proofreadingIterator == null) {
        return null;
      }
      return UnoRuntime.queryInterface(XProofreadingIterator.class, proofreadingIterator);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return null;           // Return null as method failed
    }
  }

  /**
   * Get the menu bar of LO/OO
   * Returns null if it fails
   */
  public static XMenuBar getMenuBar(XComponent xComponent) {
    try {
      XModel xModel = UnoRuntime.queryInterface(XModel.class, xComponent);
      if (xModel == null) {
        WtMessageHandler.printToLogFile("SingleDocument: setDokumentListener: XModel not found!");
        return null;
      }
      XController xController = xModel.getCurrentController();
      if (xController == null) {
        WtMessageHandler.printToLogFile("SingleDocument: setDokumentListener: XController not found!");
        return null;
      }
      XFrame frame = xController.getFrame();
      if (frame == null) {
        return null;
      }
      XPropertySet propSet = UnoRuntime.queryInterface(XPropertySet.class, frame);
      if (propSet == null) {
        return null;
      }
      XLayoutManager layoutManager = UnoRuntime.queryInterface(XLayoutManager.class,  propSet.getPropertyValue("LayoutManager"));
      if (layoutManager == null) {
        return null;
      }
      XUIElement oMenuBar = layoutManager.getElement(MENU_BAR); 
      XPropertySet props = UnoRuntime.queryInterface(XPropertySet.class, oMenuBar); 
      return UnoRuntime.queryInterface(XMenuBar.class,  props.getPropertyValue("XMenuBar"));
    } catch (DisposedException e) {
      WtMessageHandler.printToLogFile("WtOfficeTools: Document is disposed");
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    }
    return null;
  }
  
  /**
   * Returns a empty Popup Menu 
   * Returns null if it fails
   */
  @Nullable
  public static XPopupMenu getPopupMenu(XComponentContext xContext) {
    try {
      if (xContext == null) {
        return null;
      }
      XMultiComponentFactory xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class,
              xContext.getServiceManager());
      if (xMCF == null) {
        return null;
      }
      Object oPopupMenu = xMCF.createInstanceWithContext("com.sun.star.awt.PopupMenu", xContext);
      if (oPopupMenu == null) {
        return null;
      }
      return UnoRuntime.queryInterface(XPopupMenu.class, oPopupMenu);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return null;           // Return null as method failed
    }
  }
  
  /**
   *  dispatch an internal LO/OO command
   */
  public static boolean dispatchCmd(String cmd, XComponentContext xContext) {
    return dispatchCmd(cmd, new PropertyValue[0], xContext);
  } 

  /**
   *  dispatch an internal LO/OO command
   *  cmd does not include the ".uno:" substring; e.g. pass "Zoom" not ".uno:Zoom"
   */
  public static boolean dispatchUnoCmd(String cmd, XComponentContext xContext) {
    return dispatchCmd((".uno:" + cmd), new PropertyValue[0], xContext);
  } 

  /**
   * Dispatch a internal LO/OO command
   */
  public static boolean dispatchCmd(String cmd, PropertyValue[] props, XComponentContext xContext) {
    try {
      if (xContext == null) {
        WtMessageHandler.printToLogFile("OfficeTools: dispatchCmd: xContext == null");
        return false;
      }
      XMultiComponentFactory xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class,
              xContext.getServiceManager());
      if (xMCF == null) {
        WtMessageHandler.printToLogFile("OfficeTools: dispatchCmd: xMCF == null");
        return false;
      }
      Object desktop = xMCF.createInstanceWithContext("com.sun.star.frame.Desktop", xContext);
      if (desktop == null) {
        WtMessageHandler.printToLogFile("OfficeTools: dispatchCmd: desktop == null");
        return false;
      }
      XDesktop xdesktop = UnoRuntime.queryInterface(XDesktop.class, desktop);
      if (xdesktop == null) {
        WtMessageHandler.printToLogFile("OfficeTools: dispatchCmd: xdesktop == null");
        return false;
      }
      Object helper = xMCF.createInstanceWithContext("com.sun.star.frame.DispatchHelper", xContext);
      if (helper == null) {
        WtMessageHandler.printToLogFile("OfficeTools: dispatchCmd: helper == null");
        return false;
      }
      XDispatchHelper dispatchHelper = UnoRuntime.queryInterface(XDispatchHelper.class, helper);
      if (dispatchHelper == null) {
        WtMessageHandler.printToLogFile("OfficeTools: dispatchCmd: dispatchHelper == null");
        return false;
      }
      XDispatchProvider provider = UnoRuntime.queryInterface(XDispatchProvider.class, xdesktop.getCurrentFrame());
      if (provider == null) {
        WtMessageHandler.printToLogFile("OfficeTools: dispatchCmd: provider == null");
        return false;
      }
      dispatchHelper.executeDispatch(provider, cmd, "", 0, props);
      return true;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return false;
    }
  }
  
  /**
   *  Get a String from local
   */
  public static String localeToString(Locale locale) {
    if (locale == null) {
      return null;
    }
    return locale.Language + (locale.Country.isEmpty() ? "" : "-" + locale.Country) + (locale.Variant.isEmpty() ? "" : "-" + locale.Variant);
  }

  /**
   *  return true if two locales are equal  
   */
  public static boolean isEqualLocale(Locale locale1, Locale locale2) {
    return (locale1.Language.equals(locale2.Language) && locale1.Country.equals(locale2.Country) 
        && locale1.Variant.equals(locale2.Variant));
  }

  /**
   *  return true if the list of locales contains the locale
   */
  public static boolean containsLocale(List<Locale> locales, Locale locale) {
    for (Locale loc : locales) {
      if (isEqualLocale(loc, locale)) {
        return true;
      }
    }
    return false;
  }
  
  /**
   * Rename old LanguageTool files and directories
   */
  public static void renameOldLtFiles() {
    try {
      String ltname = "LanguageTool";
      File baseDir = getBaseDir();
      if (baseDir == null || !baseDir.isDirectory()) {
        return;
      }
      File wtDir = null;
      File oldLtDir = null;
      if (SystemUtils.IS_OS_WINDOWS) {
        String parent = baseDir.getParent();
        File oldDir = new File(parent);
        oldDir = new File(oldDir, "languagetool.org");
        if (!oldDir.exists()) {
          return;
        }
        oldLtDir = new File(oldDir, ltname);
      } else {
        oldLtDir = new File(baseDir, ltname);
      }
      if (!oldLtDir.exists()) {
        return;
      }
      File lDir = new File(oldLtDir, "LibreOffice");
      File oDir = new File(oldLtDir, "OpenOffice");
      if (!lDir.exists() && ! oDir.exists()) {
        return;
      }
      wtDir = new File(baseDir, APPLICATION_ID);
      if (wtDir.exists()) {
        File oCfg = new File(wtDir, "LibreOffice/Languagetool.cfg");
        File cfg = new File(wtDir, "LibreOffice/" + CONFIG_FILE);
        if (cfg.exists() || oCfg.exists()) {
          return;
        }
        oCfg = new File(wtDir, "OpenOffice/Languagetool-ooo.cfg");
        cfg = new File(wtDir, "OpenOffice/" + OOO_CONFIG_FILE);
        if (cfg.exists() || oCfg.exists()) {
          return;
        }
        File tmpDir = new File(wtDir.getAbsolutePath());
        File newDir = new File(baseDir, APPLICATION_ID + ".sv");
        tmpDir.renameTo(newDir);
      }
      oldLtDir.renameTo(wtDir);
      File newDir = wtDir;
      wtDir = new File(newDir, "LibreOffice");
      if (wtDir.exists()) {
        renameOldFile(wtDir, "Languagetool.cfg", CONFIG_FILE);
        renameOldFile(wtDir, "LanguageTool.log", LOG_FILE_SP);
        renameOldFile(wtDir, "LanguageToolSpell.log", LOG_FILE_SP);
        renameOldFile(wtDir, "LT_AI_Instructions.dat", "WT_AI_Instructions.dat");
        renameOldFile(wtDir, "LT_Statistical_Analyzes.cfg", "WT_Statistical_Analyzes.cfg");
        File cacheDir = new File(wtDir, "cache");
        if (cacheDir.exists()) {
          File[] files = cacheDir.listFiles();
          for (File f : files) {
            f.delete();
          }
        }
      }
      wtDir = new File(newDir, "OpenOffice");
      if (wtDir.exists()) {
        renameOldFile(wtDir, "Languagetool-ooo.cfg", OOO_CONFIG_FILE);
        renameOldFile(wtDir, "LanguageTool.log", LOG_FILE_SP);
        File cacheDir = new File(wtDir, "cache");
        if (cacheDir.exists()) {
          File[] files = cacheDir.listFiles();
          for (File f : files) {
            f.delete();
          }
        }
      }
    } finally {
    }
    
  }
  
  private static void renameOldFile(File dir, String oldName, String newName) {
    File oldFile = new File(dir, oldName);
    if (oldFile.exists()) {
      File newFile = new File(dir, newName);
      oldFile.renameTo(newFile);
    }
  }

  /**
   * Returns old configuration file
   *//*
  public static File getOldConfigFile() {
    String homeDir = System.getProperty("user.home");
    if (homeDir == null) {
      WtMessageHandler.showError(new RuntimeException("Could not get home directory"));
      return null;
    }
    return new File(homeDir, OLD_CONFIG_FILE);
  }
*/
  
  /**
   * Get system depend base directory
   */
  private static File getBaseDir() {
    String userHome = null;
    File directory;
    try {
      userHome = System.getProperty("user.home");
    } catch (SecurityException ex) {
    }
    if (userHome == null) {
      WtMessageHandler.showError(new RuntimeException("Could not get home directory"));
      directory = null;
    } else if (SystemUtils.IS_OS_WINDOWS) {
      // Path: \\user\<YourUserName>\AppData\Roaming\writingtool.org\WritingTool\LibreOffice  
      File appDataDir = null;
      try {
        String appData = System.getenv("APPDATA");
        if (!StringUtils.isEmpty(appData)) {
          appDataDir = new File(appData);
        }
      } catch (SecurityException ex) {
      }
      if (appDataDir != null && appDataDir.isDirectory()) {
        String path = VENDOR_ID + "\\";
        directory = new File(appDataDir, path);
      } else {
        String path = "Application Data\\" + VENDOR_ID + "\\";
        directory = new File(userHome, path);
      }
    } else if (SystemUtils.IS_OS_LINUX) {
      // Path: /home/<YourUserName>/.config/WritingTool/LibreOffice  
      File appDataDir = null;
      try {
        String xdgConfigHome = System.getenv("XDG_CONFIG_HOME");
        if (!StringUtils.isEmpty(xdgConfigHome)) {
          appDataDir = new File(xdgConfigHome);
          if (!appDataDir.isAbsolute()) {
            //https://specifications.freedesktop.org/basedir-spec/basedir-spec-latest.html
            //All paths set in these environment variables must be absolute.
            //If an implementation encounters a relative path in any of these
            //variables it should consider the path invalid and ignore it.
            appDataDir = null;
          }
        }
      } catch (SecurityException ex) {
      }
      if (appDataDir != null && appDataDir.isDirectory()) {
        directory = appDataDir;
      } else {
        String path = ".config/";
        directory = new File(userHome, path);
      }
    } else if (SystemUtils.IS_OS_MAC_OSX) {
      String path = "Library/Application Support/";
      directory = new File(userHome, path);
    } else {
      String path = ".";
      directory = new File(userHome, path);
    }
    return directory;
  }
  
  /**
   * Returns directory to store every information for LT office extension
   * @since 4.7
   */
  public static File getWtConfigDir() {
    return getWtConfigDir(null);
  }

  public static File getWtConfigDir(XComponentContext xContext) {
    if (OFFICE_EXTENSION_ID == null) {
      if (WtVersionInfo.ooName == null && xContext == null) {
        OFFICE_EXTENSION_ID = "LibreOffice";
      } else {
        if (WtVersionInfo.ooName == null) {
          WtVersionInfo.init(xContext);
        }
        OFFICE_EXTENSION_ID = WtVersionInfo.ooName;
      }
    }
    File directory = getBaseDir();
    String path;
    if (SystemUtils.IS_OS_WINDOWS || SystemUtils.IS_OS_LINUX || SystemUtils.IS_OS_MAC_OSX) {
      path = APPLICATION_ID + "/" + OFFICE_EXTENSION_ID + "/";
    } else {
      path = "." + APPLICATION_ID + "/" + OFFICE_EXTENSION_ID + "/";
    }
    directory = new File(directory, path);
/*    
    while (atWork) {
      try {
        Thread.sleep(10);
      } catch (InterruptedException e) {
      }
    }
*/    
    if (directory != null && !directory.exists()) {
      directory.mkdirs();
    }
    return directory;
  }
  
  /**
   * Returns log file 
   */
  public static String getLogFilePath(boolean isSpellchecker) {
    return getLogFilePath(null, isSpellchecker);
  }

  public static String getLogFilePath(XComponentContext xContext, boolean isSpellchecker) {
    if (isSpellchecker) {
      return new File(getWtConfigDir(xContext), LOG_FILE_SP).getAbsolutePath();
    }
    return new File(getWtConfigDir(xContext), LOG_FILE).getAbsolutePath();
  }
  
  /**
   * Returns statistical analyzes configuration file 
   */
  public static String getStatisticalConfigFilePath() {
    return new File(getWtConfigDir(), STATISTICAL_ANALYZES_CONFIG_FILE).getAbsolutePath();
  }

  /**
   * Returns directory to saves caches
   * @since 5.2
   */
  public static File getCacheDir() {
    return getCacheDir(null);
  }
  
  public static File getCacheDir(XComponentContext xContext) {
    File cacheDir = new File(getWtConfigDir(xContext), CACHE_ID);
    if (cacheDir != null && !cacheDir.exists()) {
      cacheDir.mkdirs();
    }
    return cacheDir;
  }
  
  public static double getMaxHeapSpace() {
    if(MAX_HEAP_SPACE < 0) {
      MAX_HEAP_SPACE = Runtime.getRuntime().maxMemory();
    }
    return MAX_HEAP_SPACE;
  }
  
  private static double getHeapLimit(double maxHeap) {
    if(LT_HEAP_LIMIT < 0) {
      LT_HEAP_LIMIT = maxHeap * LT_HEAP_LIMIT_FACTOR;
    }
    return LT_HEAP_LIMIT;
  }
  
  public static double getCurrentHeapRatio() {
    return (Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory()) / LT_HEAP_LIMIT;
  }
  
  public static boolean isHeapLimitReached() {
    long usedHeap = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
    return (LT_HEAP_LIMIT < usedHeap);
  }
  
  /**
   * Get information about Java as String
   */
  public static String getJavaInformation () {
    return "Java-Version: " + System.getProperty("java.version") + ", max. Heap-Space: " + ((int) (getMaxHeapSpace()/1048576)) +
        " MB, LT Heap Space Limit: " + ((int) (getHeapLimit(getMaxHeapSpace())/1048576)) + " MB";
  }
  
  /**
   * Handles files, jar entries, and deployed jar entries in a zip file (EAR).
   * @return A String with the formated date if it can be determined, or an empty string if not.
   *//*
  private static String getClassBuildTime() {
      Date date = null;
      WtOfficeTools wtTools = new WtOfficeTools();
      URL resource = wtTools.getClass().getResource(wtTools.getClass().getSimpleName() + ".class");
      if (resource != null) {
          if (resource.getProtocol().equals("file")) {
              try {
                date = new Date(new File(resource.toURI()).lastModified());
              } catch (URISyntaxException ignored) { }
          } else if (resource.getProtocol().equals("jar")) {
              String path = resource.getPath();
              date = new Date( new File(path.substring(5, path.indexOf("!"))).lastModified() );    
          } else if (resource.getProtocol().equals("zip")) {
              String path = resource.getPath();
              File jarFileOnDisk = new File(path.substring(0, path.indexOf("!")));
              try(JarFile jFile = new JarFile (jarFileOnDisk)) {
                  ZipEntry zEntry = jFile.getEntry (path.substring(path.indexOf("!") + 2));
                  long zeTimeLong = zEntry.getTime ();
                  Date zeTimeDate = new Date(zeTimeLong);
                  date = zeTimeDate;
              } catch (IOException|RuntimeException ignored) { }
          }
      }
      if (date == null) {
        return "";
      }
      OffsetDateTime dateTime = date.toInstant().atOffset(ZoneOffset.UTC);
      return dateTime.format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss Z"));
  }
*/
  /**
   * Get WritingTool Image
   */
  public static Image getLtImage() {
    try {
      URL url = WtOfficeTools.class.getResource("/images/WTSmall.png");
      return ImageIO.read(url);
    } catch (IOException e) {
      WtMessageHandler.showError(e);
    }
    return null;
  }
  
  /**
   * Get WritingTool Image
   */
  public static ImageIcon getLtImageIcon(boolean big) {
    URL url;
    if (big) {
      url = WtOfficeTools.class.getResource("/images/WTBig.png");
    } else {
      url = WtOfficeTools.class.getResource("/images/WTSmall.png");
    }
    return new ImageIcon(url);
  }

  /**
   * Get path to WT resource 
   */
  public static InputStream getWtRessourceAsInputStream(String fileName, Locale locale) {
    String resource = RESOURCE_PATH + locale.Language + "/" + fileName;
    return WtOfficeTools.class.getResourceAsStream(resource);
  }

  /**
   * Get a boolean value from an Object
   */
  public static boolean getBooleanValue(Object o) {
    if (o != null && o instanceof Boolean) {
      return ((Boolean) o).booleanValue();
    }
    return false;
  }
  
  /**
   * timestamp for last key release
   *//*
  public static void setKeyReleaseTime(long time) {
    lastKeyRelease = time;
  }
*/
  public static void waitForLO() {
    while (WtDocumentCursorTools.isBusy() || WtViewCursorTools.isBusy() || WtFlatParagraphTools.isBusy()) {
      try {
        synchronized (waitObj) {
          numLoWaits++;
          if (numLoWaits > MAX_LO_WAITS) {
            WtMessageHandler.printToLogFile("waitForLO: Wait for more than " + MAX_LO_WAITS/100 + " seconds, "
                + "DocumentCursorTools.isBusy: " + WtDocumentCursorTools.isBusy() + ", "
                + "ViewCursorTools.isBusy: " + WtViewCursorTools.isBusy() + ", "
                + "FlatParagraphTools.isBusy: " + WtFlatParagraphTools.isBusy() + ": "
                + "Free Lock and continue.");
            if (WtDocumentCursorTools.isBusy()) {
              WtDocumentCursorTools.reset();
            }
            if (WtViewCursorTools.isBusy()) {
              WtViewCursorTools.reset();
            }
            if (WtFlatParagraphTools.isBusy()) {
              WtFlatParagraphTools.reset();
            }
          }
        }
        Thread.sleep(10);
      } catch (InterruptedException e) {
        WtMessageHandler.printException(e);
      }
    }
  }
/*  
  public static void waitForLoDic() {
    long spellDiff = KEY_RELEASE_TOLERANCE - System.currentTimeMillis() + lastKeyRelease;
    while (DocumentCursorTools.isBusy() || ViewCursorTools.isBusy() || FlatParagraphTools.isBusy() || spellDiff > 0) {
      try {
        Thread.sleep(spellDiff < 10 ? 10 : spellDiff);
      } catch (InterruptedException e) {
        MessageHandler.printException(e);
      }
      spellDiff = KEY_RELEASE_TOLERANCE - System.currentTimeMillis() + lastKeyRelease;
    }
  }
*/  
  public static void waitForLtDictionary() {
    while (WtDictionary.isActivating()) {
      try {
        Thread.sleep(10);
      } catch (InterruptedException e) {
        WtMessageHandler.printException(e);
      }
    }
  }
  
  /**
   * sleep for n milliseconds
   */
  public static void sleep(int n) {
    try {
      Thread.sleep(n);
    } catch (InterruptedException e) {
      WtMessageHandler.printException(e);
    }
  }
  
  /**
   * get the LO locale from a language
   */
  public static Locale getLocalFromLanguage(Language lang) {
    String[] countries = lang.getCountries();
    String country = countries.length != 1 ? "" : countries[0];
    return new Locale(lang.getShortCode(), country, lang.getVariant() == null ? "" : lang.getVariant());
  }

  /**
   * Are statistical rules defined for this language?
   */
  public static boolean hasStatisticalStyleRules(XComponentContext xContext) {
    Language lang = getDefaultLanguage(xContext);
    if (lang == null) {
      return false;
    }
    return hasStatisticalStyleRules(lang);
  }

  public static boolean hasStatisticalStyleRules(Language lang) {
    try {
      for (Rule rule : lang.getRelevantRules(WtOfficeTools.getMessageBundle(), null, lang, new ArrayList<>())) {
        if (rule instanceof AbstractStatisticSentenceStyleRule || rule instanceof AbstractStatisticStyleRule ||
            rule instanceof ReadabilityRule || rule instanceof AbstractStyleTooOftenUsedWordRule) {
          return true; 
        }
      }
    } catch (IOException e) {
    }
    return false;
  }
  
  /**
   * Gets the ResourceBundle (i18n strings) for the default language of the user's system.
   */
  public static ResourceBundle getMessageBundle() {
    return getMessageBundle(null);
  }
  
  private static String getWtMessageResource(java.util.Locale locale) {
    return RESOURCES + "." + locale.getLanguage() + ".MessagesBundle";
  }

  private static String getWtUrlResource(java.util.Locale locale) {
    return RESOURCES + "." + locale.getLanguage() + ".URLsBundle";
  }

  public static ResourceBundle getMessageBundle(Language lang) {
    java.util.Locale locale = lang == null ? java.util.Locale.getDefault() : lang.getLocaleWithCountryAndVariant();
    ResourceBundle wtBundle = null;
    ResourceBundle ltBundle = null;
    ResourceBundle wtFallbackBundle = null;
    try {
      wtBundle = ResourceBundle.getBundle(getWtMessageResource(locale), locale);
    } catch (MissingResourceException ex) {
    }
    try {
      ltBundle = JLanguageTool.getDataBroker().getResourceBundle(MESSAGE_BUNDLE, locale);
    } catch (MissingResourceException ex) {
    }
    locale = java.util.Locale.ENGLISH;
    try {
      wtFallbackBundle = ResourceBundle.getBundle(getWtMessageResource(locale), locale);
    } catch (MissingResourceException ex) {
    }
    ResourceBundle fallbackBundle = JLanguageTool.getDataBroker().getResourceBundle(MESSAGE_BUNDLE, locale);
    return new WtResourceBundle(wtBundle, wtFallbackBundle, ltBundle, fallbackBundle);
  }

  /**
   * Gets the URL from URLResourceBundle (i18n strings) for the default language of the user's system.
   */
  public static String getUrl(String key) {
    return getUrl(null, key);
  }

  public static String getUrl(Language lang, String key) {
    java.util.Locale locale = lang == null ? java.util.Locale.getDefault() : lang.getLocaleWithCountryAndVariant();
    ResourceBundle bundle = null;
    try {
      bundle = ResourceBundle.getBundle(getWtUrlResource(locale), locale);
      return WT_SERVER_URL + "/" + bundle.getString(key);
    } catch (MissingResourceException ex) {
    }
    locale = java.util.Locale.ENGLISH;
    bundle = ResourceBundle.getBundle(getWtUrlResource(locale), locale);
    return WT_SERVER_URL + "/" + bundle.getString(key);
  }
  
/**
 * convert array of WtProofreadingError to array of SingleProofreadingError
 */
  public static SingleProofreadingError[] wtErrorsToProofreading(WtProofreadingError[] errors) {
    SingleProofreadingError[] sErrors = new SingleProofreadingError[errors.length];
    for (int i = 0; i < errors.length; i++) {
      sErrors[i] = errors[i].toSingleProofreadingError();
    }
    return sErrors;
  }
  
/**
 * convert array of SingleProofreadingError to array of WtProofreadingError
 */
  public static WtProofreadingError[] proofreadingToWtErrors(SingleProofreadingError[] errors) {
    WtProofreadingError[] wErrors = new WtProofreadingError[errors.length];
    for (int i = 0; i < errors.length; i++) {
      wErrors[i] = new WtProofreadingError(errors[i]);
    }
    return wErrors;
  }

  /**
   * get formated WT header
   */
  public static String getFormatedWtHeader(ResourceBundle messages) {
    return "<html><FONT SIZE=\"+2\"><b>"
        + WtOfficeTools.WT_NAME + " - " 
        + messages.getString("loAboutLtDesc") + "</b></FONT></html>";
  }

  /**
   * get formated license information
   */
  public static String getFormatedLicenseInformation() {
    return "<html>"
        + "<p>Copyright (C) 2024 Fred Kruse - "
        + "<a href=\"" + WtOfficeTools.WT_SERVER + "\">" + WtOfficeTools.WT_SERVER + "</a><br>  <br>"
        + "based on LanguageTool - "
        + "Copyright (C) 2005-2024 the LanguageTool community and Daniel Naber.<br>  <br>"
        + "WritingTool and LanguageTool are licensed under the GNU Lesser General Public License.<br>"
        + "</html>";
  }

  /**
   * get HTML formated version information
   */
  public static String getFormatedHtmlVersionInformation() {
    return String.format("<html>"
        + "<p>WritingTool %s (%s)<br>"
        + "based on LanguageTool %s (%s, %s)<br>"
        + "OS: %s %s (%s)<br>"
        + "%s %s%s (%s), %s<br>"
        + "Java version: %s (%s)<br>"
        + "Java max/total/free memory: %sMB, %sMB, %sMB</p>"
        + "</html>", 
         WtVersionInfo.wtVersion,
         WtVersionInfo.wtBuildDate,
         WtVersionInfo.ltVersion(),
         WtVersionInfo.ltBuildDate(),
         WtVersionInfo.ltShortGitId(),
         System.getProperty("os.name"),
         System.getProperty("os.version"),
         System.getProperty("os.arch"),
         WtVersionInfo.ooName,
         WtVersionInfo.ooVersion,
         WtVersionInfo.ooExtension,
         WtVersionInfo.ooVendor,
         WtVersionInfo.ooLocale,
         WtVersionInfo.javaVersion,
         WtVersionInfo.javaVendor,
         Runtime.getRuntime().maxMemory()/1024/1024,
         Runtime.getRuntime().totalMemory()/1024/1024,
         Runtime.getRuntime().freeMemory()/1024/1024);
  }

  /**
   * get flat formated version information
   */
  public static String getFormatedTextVersionInformation() {
    return String.format("\nWritingTool %s (%s)\n"
        + "based on LanguageTool %s (%s, %s)\n"
        + "OS: %s %s (%s)\n"
        + "%s %s%s (%s), %s\n"
        + "Java version: %s (%s)\n"
        + "Java max/total/free memory: %sMB, %sMB, %sMB\n",
         WtVersionInfo.wtVersion,
         WtVersionInfo.wtBuildDate,
         WtVersionInfo.ltVersion(),
         WtVersionInfo.ltBuildDate(),
         WtVersionInfo.ltShortGitId(),
         System.getProperty("os.name"),
         System.getProperty("os.version"),
         System.getProperty("os.arch"),
         WtVersionInfo.ooName,
         WtVersionInfo.ooVersion,
         WtVersionInfo.ooExtension,
         WtVersionInfo.ooVendor,
         WtVersionInfo.ooLocale,
         System.getProperty("java.version"),
         System.getProperty("java.vm.vendor"),
         Runtime.getRuntime().maxMemory()/1024/1024,
         Runtime.getRuntime().totalMemory()/1024/1024,
         Runtime.getRuntime().freeMemory()/1024/1024);
  }

  /**
   * get formated extension maintainer
   */
  public static String getFormatedExtensionMaintainer() {
    return String.format("<html>"
        + "<p>Maintainer of the office extension: %s</p>"
        + "<p>Maintainers or former maintainers of the language modules -<br>"
        + "(*) means language is unmaintained in LanguageTool:</p><br>"
        + "</html>", WtOfficeTools.EXTENSION_MAINTAINER);
  }

  /**
   * get formated LanguageTool maintainer
   */
  public static String getFormatedLanguageToolMaintainers(ResourceBundle messages) {
    TreeMap<String, Language> list = new TreeMap<>();
    for (Language lang : Languages.get()) {
      if (!lang.isVariant()) {
        if (lang.getMaintainers() != null) {
          list.put(messages.getString(lang.getShortCode()), lang);
        }
      }
    }
    StringBuilder str = new StringBuilder();
    str.append("<table border=0 cellspacing=0 cellpadding=0>");
    for (Map.Entry<String, Language> entry : list.entrySet()) {
      str.append("<tr valign=\"top\"><td>");
      str.append(entry.getKey());
      if (entry.getValue().getMaintainedState() == LanguageMaintainedState.LookingForNewMaintainer) {
        str.append("(*)");
      }
      str.append(":</td>");
      str.append("<td>&nbsp;</td>");
      str.append("<td>");
      int i = 0;
      Contributor[] maintainers = list.get(entry.getKey()).getMaintainers();
      if (maintainers != null) {
        for (Contributor contributor : maintainers) {
          if (i > 0) {
            str.append(", ");
            if (i % 3 == 0) {
              str.append("<br>");
            }
          }
          str.append(contributor.getName());
          i++;
        }
      }
      str.append("</td></tr>");
    }
    str.append("</table>");
    return str.toString();
  }

  /**
   * Handle logLevel for debugging and development
   */
  public static void setLogLevel(String logLevel) {
    if (logLevel != null) {
      String[] levels = logLevel.split(LOG_DELIMITER);
      for (String level : levels) {
        if (level.equals("1") || level.equals("2") || level.equals("3") || level.equals("4") || 
            level.equals("5") || level.equals("6") || level.startsWith("all:")) {
          int numLevel;
          if (level.startsWith("all:")) {
            String[] levelAll = level.split(":");
            if (levelAll.length != 2) {
              continue;
            }
            numLevel = Integer.parseInt(levelAll[1]);
          } else {
            numLevel = Integer.parseInt(level);
          }
          if (numLevel > 0) {
            DEBUG_MODE_MD = true;
            DEBUG_MODE_TQ = true;
            if (DEBUG_MODE_SD == 0) {
              DEBUG_MODE_SD = numLevel;
            }
            if (DEBUG_MODE_SC == 0) {
              DEBUG_MODE_SC = numLevel;
            }
            if (DEBUG_MODE_CR == 0) {
              DEBUG_MODE_CR = numLevel;
            }
            if (DEBUG_MODE_AI == 0) {
              DEBUG_MODE_AI = numLevel;
            }
          }
          if (numLevel > 1) {
            DEBUG_MODE_DC = true;
            DEBUG_MODE_LM = true;
          }
          if (numLevel > 2) {
            DEBUG_MODE_FP = true;
          }
        } else if (level.startsWith("sd:") || level.startsWith("sc:") 
            || level.startsWith("cr:") || level.startsWith("ai:")) {
          String[] levelSD = level.split(":");
          if (levelSD.length != 2) {
            continue;
          }
          int numLevel = Integer.parseInt(levelSD[1]);
          if (numLevel > 0) {
            if (levelSD[0].equals("sd")) {
              DEBUG_MODE_SD = numLevel;
            } else if (levelSD[0].equals("sc")) {
              DEBUG_MODE_SC = numLevel;
            } else if (levelSD[0].equals("cr")) {
              DEBUG_MODE_CR = numLevel;
            } else if (levelSD[0].equals("ai")) {
              DEBUG_MODE_AI = numLevel;
            }
          }
        } else if (level.equals("md")) {
          DEBUG_MODE_MD = true;
        } else if (level.equals("dc")) {
          DEBUG_MODE_DC = true;
        } else if (level.equals("fp")) {
          DEBUG_MODE_FP = true;
        } else if (level.equals("lm")) {
          DEBUG_MODE_LM = true;
        } else if (level.equals("tq")) {
          DEBUG_MODE_TQ = true;
        } else if (level.equals("ld")) {
          DEBUG_MODE_LD = true;
        } else if (level.equals("cd")) {
          DEBUG_MODE_CD = true;
        } else if (level.equals("io")) {
          DEBUG_MODE_IO = true;
        } else if (level.equals("rm")) {
          DEBUG_MODE_RM = true;
        } else if (level.equals("sr")) {
          DEBUG_MODE_SR = true;
        } else if (level.equals("sp")) {
          DEBUG_MODE_SP = true;
        } else if (level.equals("ta")) {
          DEBUG_MODE_TA = true;
        } else if (level.startsWith("tm")) {
          String[] levelTm = level.split(":");
          if (levelTm[0].equals("tm")) {
            DEBUG_MODE_TM = true;
            if(levelTm.length > 1) {
              int time = Integer.parseInt(levelTm[1]);
              if (time >= 0) {
                TIME_TOLERANCE = time;
              }
            }
          }
        } else if (level.equals("st")) {
          DEVELOP_MODE_ST = true;
        } else if (level.equals("dev")) {
          DEVELOP_MODE = true;
        }
      }
    }
  }
}
