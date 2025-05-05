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

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.jetbrains.annotations.Nullable;
import org.writingtool.WtProofreadingError;
import org.writingtool.WtPropertyValue;
import org.writingtool.WtSingleCheck.SentenceErrors;

import com.sun.star.beans.Property;
import com.sun.star.beans.PropertyState;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.container.XStringKeyMap;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.Locale;
import com.sun.star.lang.XComponent;
import com.sun.star.text.TextMarkupDescriptor;
import com.sun.star.text.TextMarkupType;
import com.sun.star.text.XFlatParagraph;
import com.sun.star.text.XFlatParagraphIterator;
import com.sun.star.text.XFlatParagraphIteratorProvider;
import com.sun.star.text.XMultiTextMarkup;
import com.sun.star.uno.UnoRuntime;

/**
 * Information about Paragraphs of LibreOffice/OpenOffice documents
 * on the basis of the LO/OO FlatParagraphs
 * @since 1.0
 * @author Fred Kruse
 */
public class WtFlatParagraphTools {
  
  /**
   * NOTE: The flatparagraph iterator doesn't work for OpenOffice (OO).
   *       So no support of cache is possible for OO
   */
  
  private static boolean debugMode; //  should be false except for testing
  
  private static int isBusy = 0;

  private XFlatParagraphIterator xFlatParaIter;
  private XFlatParagraph lastFlatPara;
  private XComponent xComponent;
  
  public WtFlatParagraphTools(XComponent xComponent) {
    debugMode = WtOfficeTools.DEBUG_MODE_FP;
    this.xComponent = xComponent;
    xFlatParaIter = getXFlatParagraphIterator(xComponent);
    lastFlatPara = getCurrentFlatParagraph();
  }
  
  /**
   * document is disposed: set all class variables to null
   */
  public void setDisposed() {
    xFlatParaIter = null;
    lastFlatPara = null;
    xComponent = null;
  }
  
  /**
   * document is disposed
   */
  public boolean isDisposed() {
    return xComponent == null;
  }

  /**
   * is valid initialization of FlatParagraphTools
   */
  public boolean isValid() {
    return xFlatParaIter != null;
  }

  /**
   * Returns XFlatParagraphIterator 
   * Returns null if it fails
   */
  @Nullable
  private XFlatParagraphIterator getXFlatParagraphIterator(XComponent xComponent) {
    isBusy++;
    try {
      if (xComponent == null) {
        return null;
      }
      XFlatParagraphIteratorProvider xFlatParaItPro 
          = UnoRuntime.queryInterface(XFlatParagraphIteratorProvider.class, xComponent);
      if (xFlatParaItPro == null) {
        return null;
      }
      return xFlatParaItPro.getFlatParagraphIterator(TextMarkupType.PROOFREADING, true);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return null;           // Return null as method failed
    } finally {
      isBusy--;
    }
  }
  
  /**
   * Initialize XFlatParagraphIterator
   * Set the new iterator only if it is not null
   */
  public void init() {
    isBusy++;
    try {
      XFlatParagraphIterator tmpFlatParaIter = getXFlatParagraphIterator(xComponent);
      if (tmpFlatParaIter != null) {
        xFlatParaIter = tmpFlatParaIter;
      }
    } finally {
      isBusy--;
    }
  }
  
  /**
   * Returns current FlatParagraph
   * Set lastFlatPara if current FlatParagraph is not null
   * Returns null if it fails
   */
  @Nullable
  private XFlatParagraph getCurrentFlatParagraph() {
    isBusy++;
    try {
      if (xFlatParaIter == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: getCurrentFlatParagraph: FlatParagraphIterator == null");
        }
        return null;
      }
      XFlatParagraph tmpFlatPara = xFlatParaIter.getNextPara();
      if (tmpFlatPara != null) {
        lastFlatPara = tmpFlatPara;
      }
      return tmpFlatPara;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return null;           // Return null as method failed
    } finally {
      isBusy--;
    }
  }
    
  /**
   * Returns last FlatParagraph not null
   * Set lastFlatPara if current FlatParagraph is not null
   * Returns null if it fails
   */
  @Nullable
  private XFlatParagraph getLastFlatParagraph() {
    getCurrentFlatParagraph();
    return lastFlatPara;
  }
    
  /**
   * is true if FlatParagraph is from Automatic Iteration
   */
  public boolean isFlatParaFromIter() {
    return (getCurrentFlatParagraph() != null);
  }

  /**
   * Change text of flat paragraph nPara 
   * delete characters between nStart and nStart + nLen, insert newText at nStart
   */
  public XFlatParagraph getFlatParagraphAt(int nPara) {
    if (isDisposed()) {
      return null;
    }
    WtOfficeTools.waitForLO();
    isBusy++;
    try {
      XFlatParagraph xFlatPara = getLastFlatParagraph();
      if (xFlatPara == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: getFlatParagraphAt: FlatParagraph == null");
        }
        return null;
      }
      XFlatParagraph tmpFlatPara = xFlatPara;
      while (tmpFlatPara != null) {
        xFlatPara = tmpFlatPara;
        tmpFlatPara = xFlatParaIter.getParaBefore(tmpFlatPara);
      }
      int num = 0;
      while (xFlatPara != null && num < nPara) {
        xFlatPara = xFlatParaIter.getParaAfter(xFlatPara);
        num++;
      }
      if (xFlatPara == null) {
        WtMessageHandler.printToLogFile("FlatParagraphTools: getFlatParagraphAt: FlatParagraph == null; n = " + num + "; nPara = " + nPara);
        return null;
      }
      return xFlatPara;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return null;             // Return null as method failed
    } finally {
      isBusy--;
    }
  }
  
  /**
   * return text of current paragraph
   * return null if it fails
   */
  public String getCurrentParaText() {
    isBusy++;
    try {
      XFlatParagraph xFlatPara = getCurrentFlatParagraph();
      if (xFlatPara == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: getCurrentParaText: FlatParagraph == null");
        }
        return null;
      }
      return new String(xFlatPara.getText());
    } finally {
      isBusy--;
    }
  }

  /**
   * Returns Current Paragraph Number from FlatParagaph
   * Returns -1 if it fails
   */
  public int getCurNumFlatParagraph() {
    isBusy++;
    try {
      XFlatParagraph xFlatPara = getCurrentFlatParagraph();
      if (xFlatPara == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: getCurNumFlatParagraph: FlatParagraph == null");
        }
        return -1;
      }
      int pos = -1;
      XFlatParagraph tmpXFlatPara = xFlatPara;
      while (tmpXFlatPara != null) {
        tmpXFlatPara = xFlatParaIter.getParaBefore(tmpXFlatPara);
        pos++;
      }
      xFlatParaIter = getXFlatParagraphIterator(xComponent);
      return pos;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return -1;           // Return -1 as method failed
    } finally {
      isBusy--;
    }
  }

  /**
   * Returns Text of all FlatParagraphs of Document
   * Returns null if it fails
   */
  @Nullable
  public FlatParagraphContainer getAllFlatParagraphs(Locale fixedLocale) {
    if (isDisposed()) {
      return null;
    }
    WtOfficeTools.waitForLO();
    isBusy++;
    try {
      XFlatParagraph xFlatPara = getLastFlatParagraph();
      if (xFlatPara == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: getAllFlatParagraphs: FlatParagraph == null");
        }
        return null;
      }
      List<String> allParas = new ArrayList<>();
      List<Locale> locales = new ArrayList<>();
      List<int[]> footnotePositions = new ArrayList<>();
      XFlatParagraph tmpFlatPara = xFlatPara;
      List<Integer> sortedTextIds = getIntPropertyValue("SortedTextId", tmpFlatPara) == -1 ? null : new ArrayList<>();
      int documentElementsCount = sortedTextIds == null ? -1 : getIntPropertyValue("DocumentElementsCount", tmpFlatPara);
      Locale locale = null;
      while (tmpFlatPara != null) {
        String text = new String(tmpFlatPara.getText());
        int len = text.length();
        allParas.add(0, text);
        footnotePositions.add(0, getIntArrayPropertyValue("FootnotePositions", tmpFlatPara));
        // add just one local for the whole paragraph
        locale = getPrimaryParagraphLanguage(tmpFlatPara, 0, len, fixedLocale, locale, false);
        locales.add(0, locale);
        if (sortedTextIds != null) {
          sortedTextIds.add(0, getIntPropertyValue("SortedTextId", tmpFlatPara));
        }
        tmpFlatPara = xFlatParaIter.getParaBefore(tmpFlatPara);
      }
      tmpFlatPara = xFlatParaIter.getParaAfter(xFlatPara);
      while (tmpFlatPara != null) {
        String text = new String(tmpFlatPara.getText());
        int len = text.length();
        allParas.add(text);
        footnotePositions.add(getIntArrayPropertyValue("FootnotePositions", tmpFlatPara));
        locale = getPrimaryParagraphLanguage(tmpFlatPara, 0, len, fixedLocale, locale, false);
        locales.add(locale);
        if (debugMode) {
          printPropertyValueInfo(tmpFlatPara);
        }
        if (sortedTextIds != null) {
          sortedTextIds.add(getIntPropertyValue("SortedTextId", tmpFlatPara));
        }
        tmpFlatPara = xFlatParaIter.getParaAfter(tmpFlatPara);
      }
      xFlatParaIter = getXFlatParagraphIterator(xComponent);
      return new FlatParagraphContainer(allParas, locales, footnotePositions, sortedTextIds, documentElementsCount);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return null;           // Return null as method failed
    } finally {
      isBusy--;
    }
  }
  
  /**
   * Returns Text of some FlatParagraphs defined in a List
   * Returns null if it fails
   */
  @Nullable
  public List<String> getFlatParagraphs(List<Integer> nParas) {
    if (isDisposed()) {
      return null;
    }
    WtOfficeTools.waitForLO();
    isBusy++;
    try {
      XFlatParagraph xFlatPara = getLastFlatParagraph();
      if (xFlatPara == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: getAllFlatParagraphs: FlatParagraph == null");
        }
        return null;
      }
      List<String> sParas = new ArrayList<>();
      XFlatParagraph tmpFlatPara = xFlatPara;
      while (tmpFlatPara != null) {
        xFlatPara = tmpFlatPara;
        tmpFlatPara = xFlatParaIter.getParaBefore(tmpFlatPara);
      }
      int nFlat = 0;
      int nPara = 0;
      while (xFlatPara != null && nPara < nParas.size()) {
        if (nFlat == nParas.get(nPara)) {
          String text = new String(xFlatPara.getText());
          sParas.add(text);
          nPara++;
        }
        xFlatPara = xFlatParaIter.getParaAfter(xFlatPara);
        nFlat++;
      }
      return sParas;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return null;           // Return null as method failed
    } finally {
      isBusy--;
    }
  }
  
  /**
   * Get a save Locale 
   */
  private static Locale getSaveLocale(String language, String country, String variant) {
    return new Locale(new String(language), new String(country), new String(variant));
  }
  
  /**
   * Get the language of paragraph 
   * @throws IllegalArgumentException 
   */
  private static Locale getParagraphLanguage(XFlatParagraph flatPara, int first, int len) throws IllegalArgumentException {
    Locale locale = flatPara.getLanguageOfText(first, len);
    if (locale == null || locale.Language.isEmpty()) {
      locale = flatPara.getPrimaryLanguageOfText(first, len);
    }
    return getSaveLocale(locale.Language, locale.Country, locale.Variant);
  }
  
  /**
   * Get language Portions of a paragraph
   * @throws IllegalArgumentException 
   */
  @SuppressWarnings("unused")
  private static List<Integer> getLanguagePortions(XFlatParagraph flatPara, int len) throws IllegalArgumentException {
    List<Integer> langPortions = new ArrayList<Integer>();
    Locale lastLocale = flatPara.getLanguageOfText(0, 1);
    for (int i = 1; i < len; i++) {
      Locale locale = flatPara.getLanguageOfText(i, 1);
      if (!WtOfficeTools.isEqualLocale(lastLocale, locale)) {
        langPortions.add(i);
        lastLocale = locale;
      }
    }
    return langPortions;
  }
  
  
  /**
   * Get the main language of paragraph 
   * @throws IllegalArgumentException 
   */
  public static Locale getPrimaryParagraphLanguage(XFlatParagraph flatPara, int start, int len, Locale fixedLocale, 
      Locale lastLocale, boolean onlyPrimary) throws Throwable {
    isBusy++;
    try {
      if (fixedLocale != null) {
        return fixedLocale;
      }
      int lenPara = flatPara.getText().length();
      if (start + len > lenPara) {
        len = lenPara - start;
      }
      if (len <= 0) {
        if (lastLocale != null) {
          return lastLocale.Variant.startsWith(WtOfficeTools.MULTILINGUAL_LABEL) 
              ? getSaveLocale(lastLocale.Language, lastLocale.Country, lastLocale.Variant.substring(WtOfficeTools.MULTILINGUAL_LABEL.length())) 
              : lastLocale;
        } else if (lenPara > 0) {
          return getParagraphLanguage(flatPara, 0, 1);
        } else {
          getSaveLocale(WtOfficeTools.IGNORE_LANGUAGE, "", "");
        }
      }
      if (len == 1) {
        return getParagraphLanguage(flatPara, start, len);
      }
      Map<Locale, Integer> locales = new HashMap<Locale, Integer>();
      for (int i = start; i < len; i++) {
        Locale locale = flatPara.getLanguageOfText(i, 1);
        boolean existingLocale = false;
        for (Locale loc : locales.keySet()) {
          if (loc.Language.equals(locale.Language)) {
            locales.put(loc, locales.get(loc) + 1);
            existingLocale = true;
            break;
          }
        }
        if (!existingLocale) {
          locales.put(locale, 1);
        }
      }
      if (locales.keySet().size() == 0) {
        if (lastLocale == null) {
          return getSaveLocale(WtOfficeTools.IGNORE_LANGUAGE, "", "");
        }
        return lastLocale.Variant.startsWith(WtOfficeTools.MULTILINGUAL_LABEL) ? 
            getSaveLocale(lastLocale.Language, lastLocale.Country, lastLocale.Variant.substring(WtOfficeTools.MULTILINGUAL_LABEL.length())) : lastLocale;
      }
      Locale biggestLocal = null;
      int biggestLocalNumber = 0;
      for (Locale loc : locales.keySet()) {
        int locNum = locales.get(loc);
        if (biggestLocal == null || locNum > biggestLocalNumber) {
          biggestLocal = loc;
          biggestLocalNumber = locNum;
        }
      }
      if (biggestLocal == null) {
        return lastLocale.Variant.startsWith(WtOfficeTools.MULTILINGUAL_LABEL) ? 
            new Locale(lastLocale.Language, lastLocale.Country, lastLocale.Variant.substring(WtOfficeTools.MULTILINGUAL_LABEL.length())) : lastLocale;
      } else if (onlyPrimary || locales.keySet().size() == 1) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: getPrimaryParagraphLanguage: locale: " + WtOfficeTools.localeToString(biggestLocal));
        }
        return biggestLocal;
      } else {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: getPrimaryParagraphLanguage: is multilingual locale: " + WtOfficeTools.localeToString(biggestLocal));
        }
        return getSaveLocale(biggestLocal.Language, biggestLocal.Country, WtOfficeTools.MULTILINGUAL_LABEL + biggestLocal.Variant);
      }
    } finally {
      isBusy--;
    }
  }

  /**
   * Get the main language of paragraph 
   * @throws IllegalArgumentException 
   */
  public Locale getPrimaryLanguageOfPartOfParagraph(int nPara, int start, int len, Locale lastLocale) throws Throwable {
    isBusy++;
    try {
      XFlatParagraph flatPara = getFlatParagraphAt(nPara);
      if (flatPara == null) {
        return lastLocale;
      }
      return getPrimaryParagraphLanguage(flatPara, start, len, null, lastLocale, true);
    } finally {
      isBusy--;
    }
  }
  
  /**
   * Get language of character
   * @throws IllegalArgumentException 
   */
  public Locale getLanguageOfWord(int nPara, int start, int len, Locale paragraphLocale) throws Throwable {
    try {
      XFlatParagraph flatPara = getFlatParagraphAt(nPara);
      isBusy++;
      if (flatPara == null || start + len > flatPara.getText().length()) {
        return paragraphLocale;
      }
      Locale locale = flatPara.getLanguageOfText(start, len);
      if (locale == null || locale.Language.isEmpty()) {
        return paragraphLocale;
      }
      return getSaveLocale(locale.Language, locale.Country, locale.Variant);
    } finally {
      isBusy--;
    }
  }
  
  /**
   * Returns Number of all FlatParagraphs of Document from current FlatParagraph
   * Returns negative value if it fails
   */
  public int getNumberOfAllFlatPara() {
    if (isDisposed()) {
      return -1;
    }
    WtOfficeTools.waitForLO();
    isBusy++;
    try {
      XFlatParagraph xFlatPara = getLastFlatParagraph();
      if (xFlatPara == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: getNumberOfAllFlatPara: FlatParagraph == null");
        }
        return -1;
      }
      XFlatParagraph tmpFlatPara = xFlatPara;
      int num = 0;
      while (tmpFlatPara != null) {
        tmpFlatPara = xFlatParaIter.getParaBefore(tmpFlatPara);
        num++;
      }
      tmpFlatPara = xFlatPara;
      num--;
      while (tmpFlatPara != null) {
        tmpFlatPara = xFlatParaIter.getParaAfter(tmpFlatPara);
        num++;
      }
      return num;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return -1;             // Return -1 as method failed
    } finally {
      isBusy--;
    }
  }

  /** 
   * Returns positions of properties by name 
   */
  public int[] getFootnotePosition(XFlatParagraph xFlatPara) {
    return getIntArrayPropertyValue("FootnotePositions", xFlatPara);
  }

  /** 
   * Returns positions of properties by name 
   */
  private Object getPropertyValueAsObject(String propName, XFlatParagraph xFlatPara) {
    try {
      if (xFlatPara == null || propName == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: getPropertyValueAsObject: " 
              + (xFlatPara == null ? "FlatParagraph" : "propName") + " == null");
        }
        return  null;
      }
      XPropertySet paraProps = UnoRuntime.queryInterface(XPropertySet.class, xFlatPara);
      if (paraProps == null) {
        WtMessageHandler.printToLogFile("FlatParagraphTools: getPropertyValueAsObject: XPropertySet == null");
        return  null;
      }
      return paraProps.getPropertyValue(propName);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    }
    return null;
  }
  
  /** 
   * Returns positions of properties by name 
   */
  private int[] getIntArrayPropertyValue(String propName, XFlatParagraph xFlatPara) {
    try {
      Object propertyValue = getPropertyValueAsObject(propName, xFlatPara);
      if (propertyValue == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: getIntArrayPropertyValue: propertyValue == null");
        }
        return  new int[]{};
      }
      if (propertyValue instanceof int[]) {
        return (int[]) propertyValue;
      } else {
        WtMessageHandler.printToLogFile("FlatParagraphTools: getIntArrayPropertyValue: Not of expected type int[]: " + propertyValue + ": " + propertyValue);
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    }
    return new int[]{};
  }
  
  /** 
   * Returns positions of properties by name 
   */
  private int getIntPropertyValue(String propName, XFlatParagraph xFlatPara) {
    try {
      Object propertyValue = getPropertyValueAsObject(propName, xFlatPara);
      if (propertyValue == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: getIntPropertyValue: propertyValue == null");
        }
        return  -1;
      }
      if (propertyValue instanceof Integer) {
        return (int) propertyValue;
      } else {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: getPropertyValues: Not of expected type int: " + propertyValue + ": " + propertyValue);
        }
        return -1;
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    }
    return -1;
  }
  
  /** 
   * Print property values to logfile
   */
  private void printPropertyValueInfo(XFlatParagraph xFlatPara) {
    if (xFlatPara == null) {
      WtMessageHandler.printToLogFile("FlatParagraphTools: printPropertyValues: FlatParagraph == null");
      return;
    }
    XPropertySet paraProps = UnoRuntime.queryInterface(XPropertySet.class, xFlatPara);
    if (paraProps == null) {
      WtMessageHandler.printToLogFile("XPropertySet == null");
      return;
    }
    WtMessageHandler.printToLogFile("FlatParagraphTools: Property Value Info:");
    try {
      XPropertySetInfo propertySetInfo = paraProps.getPropertySetInfo();
      
      for (Property property : propertySetInfo.getProperties()) {
        int nValue;
        if (property.Name.equals("FootnotePositions") || property.Name.equals("FieldPositions")) {
          nValue = ((int[]) paraProps.getPropertyValue(property.Name)).length;
        } else {
          nValue = (int) paraProps.getPropertyValue(property.Name);
        }
        WtMessageHandler.printToLogFile("Name : " + property.Name + "; Type : " + property.Type.getTypeName() + "; Value : " + nValue + "; Handle : " + property.Handle);
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    }
    return;
  }
  
  /**
   * Marks all paragraphs as checked with exception of the paragraphs "from" to "to"
   */
  public void setFlatParasAsChecked(int from, int to, List<Boolean> isChecked) {
    if (isDisposed()) {
      return;
    }
    WtOfficeTools.waitForLO();
    isBusy++;
    try {
      XFlatParagraph xFlatPara = getLastFlatParagraph();
      if (xFlatPara == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: setFlatParasAsChecked: FlatParagraph == null");
        }
        return;
      }
      if (isChecked == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: setFlatParasAsChecked: List isChecked == null");
        }
        isChecked  = new ArrayList<>();
      }
      XFlatParagraph tmpFlatPara = xFlatPara;
      XFlatParagraph startFlatPara = xFlatPara;
      while (tmpFlatPara != null) {
        startFlatPara = tmpFlatPara;
        tmpFlatPara = xFlatParaIter.getParaBefore(tmpFlatPara);
      }
      tmpFlatPara = startFlatPara;
      int num = 0;
      while (tmpFlatPara != null && num < from) {
        if (debugMode && num >= isChecked.size()) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: setFlatParasAsChecked: List isChecked == null");
        }
        if (num < isChecked.size() && isChecked.get(num)) {
          tmpFlatPara.setChecked(TextMarkupType.PROOFREADING, true);
        }
        tmpFlatPara = xFlatParaIter.getParaAfter(tmpFlatPara);
        num++;
      }
      while (tmpFlatPara != null && num < to) {
        tmpFlatPara.setChecked(TextMarkupType.PROOFREADING, false);
        tmpFlatPara = xFlatParaIter.getParaAfter(tmpFlatPara);
        num++;
      }
      while (tmpFlatPara != null) {
        if (debugMode && num >= isChecked.size()) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: setFlatParasAsChecked: List isChecked == null");
        }
        if (num < isChecked.size() && isChecked.get(num)) {
          tmpFlatPara.setChecked(TextMarkupType.PROOFREADING, true);
        }
        tmpFlatPara = xFlatParaIter.getParaAfter(tmpFlatPara);
        num++;
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
    } finally {
      isBusy--;
    }
  }
  
  /**
   * Marks all paragraphs as checked
   */
  public void setFlatParasAsChecked() {
    if (isDisposed()) {
      return;
    }
    WtOfficeTools.waitForLO();
    isBusy++;
    try {
      XFlatParagraph xFlatPara = getLastFlatParagraph();
      if (xFlatPara == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: setFlatParasAsChecked: FlatParagraph == null");
        }
        return;
      }
      XFlatParagraph tmpFlatPara = xFlatPara;
      while (tmpFlatPara != null) {
        tmpFlatPara.setChecked(TextMarkupType.PROOFREADING, true);
        tmpFlatPara = xFlatParaIter.getParaBefore(tmpFlatPara);
      }
      tmpFlatPara = xFlatPara;
      while (tmpFlatPara != null) {
        tmpFlatPara.setChecked(TextMarkupType.PROOFREADING, true);
        tmpFlatPara = xFlatParaIter.getParaAfter(tmpFlatPara);
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
    } finally {
      isBusy--;
    }
  }
  
  /**
   * Get information of checked status of all paragraphs
   */
  public List<Boolean> isChecked(List<Integer> changedParas, int nDiv) {
    if (isDisposed()) {
      return null;
    }
    WtOfficeTools.waitForLO();
    isBusy++;
    List<Boolean> isChecked = new ArrayList<>();
    try {
      XFlatParagraph xFlatPara = getLastFlatParagraph();
      if (xFlatPara == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: isChecked: FlatParagraph == null");
        }
        return null;
      }
      XFlatParagraph tmpFlatPara = xFlatPara;
      XFlatParagraph startFlatPara = xFlatPara;
      while (tmpFlatPara != null) {
        startFlatPara = tmpFlatPara;
        tmpFlatPara = xFlatParaIter.getParaBefore(tmpFlatPara);
      }
      tmpFlatPara = startFlatPara;
      for (int i = 0; tmpFlatPara != null; i++) {
        boolean dontCheck = (changedParas == null || !changedParas.contains(i - nDiv))
            && tmpFlatPara.isChecked(TextMarkupType.PROOFREADING);
        isChecked.add(dontCheck);
        tmpFlatPara = xFlatParaIter.getParaAfter(tmpFlatPara);
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
    } finally {
      isBusy--;
    }
    return isChecked;
  }
  
  /**
   * Set marks to changed paragraphs
   * if override is true existing marks are removed and marks are new set
   * else the marks are added to the existing marks
   */

  public void markParagraphs(Map<Integer, List<SentenceErrors>> changedParas) throws Throwable {
    isBusy++;
    try {
      if (changedParas == null || changedParas.isEmpty()) {
        return;
      }
      XFlatParagraph xFlatPara = getLastFlatParagraph();
      if (xFlatPara == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: markParagraphs: FlatParagraph == null");
        }
        return;
      }
      // treat text cursor separately because of performance reasons for big texts
      XFlatParagraph tmpFlatPara = xFlatPara;
      XFlatParagraph startFlatPara = xFlatPara;
      while (tmpFlatPara != null) {
        startFlatPara = tmpFlatPara;
        tmpFlatPara = xFlatParaIter.getParaBefore(tmpFlatPara);
      }
      tmpFlatPara = startFlatPara;
      int num = 0;
      int nMarked = 0;
      while (tmpFlatPara != null && nMarked < changedParas.size()) {
        if (changedParas.containsKey(num)) {
          addMarksToOneParagraph(tmpFlatPara, changedParas.get(num));
          if (debugMode) {
            WtMessageHandler.printToLogFile("FlatParagraphTools: mark Paragraph: " + num + ", Text: " + tmpFlatPara.getText());
          }
          nMarked++;
        }
        tmpFlatPara = xFlatParaIter.getParaAfter(tmpFlatPara);
        num++;
      }
      if (debugMode && tmpFlatPara == null) {
        WtMessageHandler.printToLogFile("FlatParagraphTools: markParagraphs: tmpFlatParagraph == null");
      }
      xFlatParaIter = getXFlatParagraphIterator(xComponent);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
    } finally {
      isBusy--;
    }
  }
  
  /**
   * add marks to existing marks of current paragraph
   */
  public void markCurrentParagraph(List<SentenceErrors> errorList) {
    isBusy++;
    try {
      if (errorList == null || errorList.size() == 0) {
        return;
      }
      XFlatParagraph xFlatPara = getCurrentFlatParagraph();
      if (xFlatPara == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: markCurrentParagraph: FlatParagraph == null");
        }
        return;
      }
      addMarksToOneParagraph(xFlatPara, errorList);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
    } finally {
      isBusy--;
    }
  }
    
  /**
   * add marks to existing marks of a paragraph
   * if override: existing marks will be overridden
   * on base of string commit
   *//*
  private void addMarksToOneParagraph(XFlatParagraph flatPara, List<SentenceErrors> errorList) throws Throwable {
    try {
      boolean isChecked = flatPara.isChecked(TextMarkupType.PROOFREADING);
      if (debugMode) {
        WtMessageHandler.printToLogFile("FlatParagraphTools: addMarksToOneParagraph: xMarkingAccess: isChecked = " + isChecked);
      }
      int paraLen = flatPara.getText().length();
      for (SentenceErrors errors : errorList) {
        XStringKeyMap props;
        for (WtProofreadingError pError : errors.sentenceErrors) {
          if (pError.nErrorStart < paraLen) {
            props = flatPara.getMarkupInfoContainer();
            WtPropertyValue[] properties = pError.aProperties;
            int color = -1;
            short type = -1;
            for (WtPropertyValue property : properties) {
              if ("LineColor".equals(property.name)) {
                color = (int) property.value;
              } else if ("LineType".equals(property.name)) {
                type = (short) property.value;
              }
            }
            if (color >= 0) {
              props.insertValue("LineColor", color);
            }
            if (type > 0) {
              props.insertValue("LineType", type);
            }
            int errLen = pError.nErrorLength;
            if (errLen + pError.nErrorStart >= paraLen) {
              errLen = paraLen - pError.nErrorStart;
            }
            if (errLen > 0) {
              flatPara.commitStringMarkup(TextMarkupType.PROOFREADING, pError.aRuleIdentifier, 
                pError.nErrorStart, errLen, props);
            }
          }
        }
  //      props = flatPara.getMarkupInfoContainer();
  //      flatPara.commitStringMarkup(TextMarkupType.SENTENCE, "Sentence", errors.sentenceStart, errors.sentenceEnd - errors.sentenceStart, props);
      }
      if (isChecked) {
        flatPara.setChecked(TextMarkupType.PROOFREADING, true);
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    }
  }
*/
  /**
   * add marks to existing marks of a paragraph
   * if override: existing marks will be overridden
   * on base of XMultiTextMarkup
   */
  private void addMarksToOneParagraph(XFlatParagraph flatPara, List<SentenceErrors> errorList) throws Throwable {
    try {
      boolean isChecked = flatPara.isChecked(TextMarkupType.PROOFREADING);
      if (debugMode) {
        WtMessageHandler.printToLogFile("FlatParagraphTools: addMarksToOneParagraph: xMarkingAccess: isChecked = " + isChecked);
      }
      int paraLen = flatPara.getText().length();
      XMultiTextMarkup xMultiTextMarkup = UnoRuntime.queryInterface(XMultiTextMarkup.class, flatPara);
      for (SentenceErrors errors : errorList) {
        XStringKeyMap props;
        TextMarkupDescriptor[] markupDescriptor = new TextMarkupDescriptor[errors.sentenceErrors.length + 1];
        int i = 0;
        for (; i < errors.sentenceErrors.length; i++) {
          WtProofreadingError pError = errors.sentenceErrors[i];
          if (pError.nErrorStart >= paraLen) {
            WtMessageHandler.printToLogFile("FlatParagraphTools: addMarksToOneParagraph: pError.nErrorStart >= paraLen: (" + 
                pError.nErrorStart + "/" + paraLen + ") : paragraph: " + flatPara.getText());
            return;
          }
          props = flatPara.getMarkupInfoContainer();
          WtPropertyValue[] properties = pError.aProperties;
          int color = -1;
          short type = -1;
          for (WtPropertyValue property : properties) {
            if ("LineColor".equals(property.name)) {
              color = (int) property.value;
            } else if ("LineType".equals(property.name)) {
              type = (short) property.value;
            }
          }
          if (color >= 0) {
            props.insertValue("LineColor", color);
          }
          if (type > 0) {
            props.insertValue("LineType", type);
          }
          int errLen = pError.nErrorLength;
          if (errLen + pError.nErrorStart >= paraLen) {
            errLen = paraLen - pError.nErrorStart;
          }
          if (errLen > 0) {
            markupDescriptor[i] = new TextMarkupDescriptor(TextMarkupType.PROOFREADING, pError.aRuleIdentifier, 
                pError.nErrorStart, errLen, props);
          }
        }
        if (i == errors.sentenceErrors.length) {
          props = flatPara.getMarkupInfoContainer();
          markupDescriptor[i] = new TextMarkupDescriptor(TextMarkupType.SENTENCE, "Sentence", 
              errors.sentenceStart, errors.sentenceEnd - errors.sentenceStart, props);
          xMultiTextMarkup.commitMultiTextMarkup(markupDescriptor);
        }
      }
      if (isChecked) {
        flatPara.setChecked(TextMarkupType.PROOFREADING, true);
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    }
  }

  /**
   * Change text of flat paragraph nPara 
   * delete characters between nStart and nStart + nLen, insert newText at nStart
   */
  public void changeTextOfParagraph(int nPara, int nStart, int nLen, String newText) throws RuntimeException {
    if (isDisposed()) {
      return;
    }
    if (newText == null) {
      throw new RuntimeException("Text is null");
    }
    WtOfficeTools.waitForLO();
    isBusy++;
    try {
      XFlatParagraph xFlatPara = getLastFlatParagraph();
      if (xFlatPara == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: changeTextOfParagraph: FlatParagraph == null");
        }
        return;
      }
      XFlatParagraph tmpFlatPara = xFlatPara;
      while (tmpFlatPara != null) {
        xFlatPara = tmpFlatPara;
        tmpFlatPara = xFlatParaIter.getParaBefore(tmpFlatPara);
      }
      int num = 0;
      while (xFlatPara != null && num < nPara) {
        xFlatPara = xFlatParaIter.getParaAfter(xFlatPara);
        num++;
      }
      if (xFlatPara == null) {
        WtMessageHandler.printToLogFile("FlatParagraphTools: changeTextOfParagraph: FlatParagraph == null; n = " + num + "; nPara = " + nPara);
        return;
      }
      int paraLen = xFlatPara.getText().length();
      if (nStart > paraLen) {
        nStart = paraLen;
      }
      if (nStart + nLen > paraLen) {
        nLen = 0;
      }
      xFlatPara.changeText(nStart, nLen, newText, new PropertyValue[0]);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return;             // Return -1 as method failed
    } finally {
      isBusy--;
    }
  }
  
  /**
   * Change text of flat paragraph nPara 
   * delete characters between nStart and nStart + nLen, insert newText at nStart
   */
  public void setLanguageOfParagraph(int nPara, int nStart, int nLen, Locale locale) throws RuntimeException {
    if (isDisposed()) {
      return;
    }
    if (locale == null) {
      throw new RuntimeException("Locale is null");
    }
    WtOfficeTools.waitForLO();
    isBusy++;
    try {
      XFlatParagraph xFlatPara = getLastFlatParagraph();
      if (xFlatPara == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: setLanguageOfParagraph: FlatParagraph == null");
        }
        return;
      }
      XFlatParagraph tmpFlatPara = xFlatPara;
      while (tmpFlatPara != null) {
        xFlatPara = tmpFlatPara;
        tmpFlatPara = xFlatParaIter.getParaBefore(tmpFlatPara);
      }
      int num = 0;
      while (xFlatPara != null && num < nPara) {
        xFlatPara = xFlatParaIter.getParaAfter(xFlatPara);
        num++;
      }
      if (xFlatPara == null) {
        WtMessageHandler.printToLogFile("FlatParagraphTools: setLanguageOfParagraph: FlatParagraph == null; n = " + num + "; nPara = " + nPara);
        return;
      }
      
      int paraLen = xFlatPara.getText().length();
      if (nStart < paraLen) {
        if (nStart + nLen > paraLen) {
          nLen = paraLen - nStart;
        }
        PropertyValue[] propertyValues = { new PropertyValue("CharLocale", -1, locale, PropertyState.DIRECT_VALUE) };
        xFlatPara.changeAttributes(nStart, nLen, propertyValues);
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return;             // Return -1 as method failed
    } finally {
      isBusy--;
    }
  }
  
  /**
   * Change text and locale of flat paragraph nPara 
   * delete characters between nStart and nStart + nLen, insert newText at nStart
   */
  public void changeTextAndLocaleOfParagraph (int nPara, int nStart, int nLen, String newText, Locale locale) throws RuntimeException {
    if (isDisposed()) {
      return;
    }
    if (newText == null) {
      throw new RuntimeException("Text is null");
    }
    if (locale == null) {
      throw new RuntimeException("Locale is null");
    }
    WtOfficeTools.waitForLO();
    isBusy++;
    try {
      XFlatParagraph xFlatPara = getLastFlatParagraph();
      if (xFlatPara == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("FlatParagraphTools: changeTextAndLocaleOfParagraph: FlatParagraph == null");
        }
        return;
      }
      XFlatParagraph tmpFlatPara = xFlatPara;
      while (tmpFlatPara != null) {
        xFlatPara = tmpFlatPara;
        tmpFlatPara = xFlatParaIter.getParaBefore(tmpFlatPara);
      }
      int num = 0;
      while (xFlatPara != null && num < nPara) {
        xFlatPara = xFlatParaIter.getParaAfter(xFlatPara);
        num++;
      }
      if (xFlatPara == null) {
        WtMessageHandler.printToLogFile("FlatParagraphTools: changeTextAndLocaleOfParagraph: FlatParagraph == null; n = " + num + "; nPara = " + nPara);
        return;
      }
      int paraLen = xFlatPara.getText().length();
      if (nStart < paraLen) {
        if (nStart + nLen > paraLen) {
          nLen = paraLen - nStart;
        }
        xFlatPara.changeText(nStart, nLen, newText, new PropertyValue[0]);
        PropertyValue[] propertyValues = { new PropertyValue("CharLocale", -1, locale, PropertyState.DIRECT_VALUE) };
        xFlatPara.changeAttributes(nStart, nLen, propertyValues);
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
      return;             // Return -1 as method failed
    } finally {
      isBusy--;
    }
  }
  
  /**
   *  Returns the status of cursor tools
   *  true: If a cursor tool in one or more threads is active
   */
  public static boolean isBusy() {
    return isBusy > 0;
  }
  
  /**
   *  Reset the busy flag
   */
  public static void reset() {
    isBusy = 0;
  }
  
  public static class FlatParagraphContainer {
    public List<String> paragraphs;
    public List<Locale> locales;
    public List<int[]> footnotePositions;
    public List<Integer> sortedTextIds;
    public int documentElementsCount;
    
    FlatParagraphContainer(List<String> paragraphs, List<Locale> locales, List<int[]> footnotePositions, 
        List<Integer> sortedTextIds, int documentElementsCount) {
      this.paragraphs = paragraphs;
      this.locales = locales;
      this.footnotePositions = footnotePositions;
      this.sortedTextIds = sortedTextIds;
      this.documentElementsCount = documentElementsCount;
    }
  }
  
}
