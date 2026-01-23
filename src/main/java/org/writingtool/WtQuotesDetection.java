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
package org.writingtool;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Pattern;

import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.i18n.LocaleDataItem;
import com.sun.star.i18n.XLocaleData;
import com.sun.star.lang.Locale;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

public class WtQuotesDetection {
  
  public static List<String> startSymbols = new ArrayList<>(Arrays.asList("“", "„", "»", "«", "\""));
  public static List<String> endSymbols   = new ArrayList<>(Arrays.asList("”", "“", "«", "»", "\""));
//  private static final List<String> singleStartSymbols = Arrays.asList("‘", "‚", "›", "‹", "'");
//  private static final List<String> singleEndSymbols   = Arrays.asList("’", "‘", "‹", "›", "'");
  private static final Pattern PUNCTUATION = Pattern.compile("[\\p{Punct}…–—&&[^\"'_]]");
  private static final Pattern PUNCT_MARKS = Pattern.compile("[\\?\\.!,]");
  
  private int nQuote = -1;
  private int nChanges = 0;
  
  private List<Integer> openingQuotes;
  private List<Integer> closingQuotes;
  
  public WtQuotesDetection() {}

  public WtQuotesDetection(XComponentContext xContext) {
    this(xContext, null);
  }

  public WtQuotesDetection(XComponentContext xContext, Locale locale) {
    try {
      setConfiguredQuotes(xContext, locale);
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    }
  }
  
  /**
   * Defines exceptions from quotes
   */
  private boolean isNotQuote (String txt, int i, int j) throws Throwable {
    if ((i == 0 || Character.isWhitespace(txt.charAt(i - 1)))
        && (i >= txt.length() - 1 || Character.isWhitespace(txt.charAt(i + 1)))) {
      return true;
    }
    if (endSymbols.get(j).equals(startSymbols.get(j))) {
      if (i > 0 && i < txt.length() - 1 
          && !Character.isWhitespace(txt.charAt(i - 1))
          && !Character.isWhitespace(txt.charAt(i + 1))) {
        String pChar = txt.substring(i - 1, i);
        String nChar = txt.substring(i + 1, i + 2);
        if(PUNCTUATION.matcher(pChar).matches()
          && !".".equals(nChar)
          && PUNCTUATION.matcher(nChar).matches()) {
          return true;
        }
      }
    }
    return false;
  }
  
  /**
   * returns the index of the opening quote
   * -1 else
   */
  private int isOpeningQuote(String txt, int i) throws Throwable {
    String tChar = txt.substring(i, i + 1);
    for (int j = 0; j < startSymbols.size(); j++) {
      if (startSymbols.get(j).equals(tChar)) {
        if (isNotQuote (txt, i, j)) {
          return -1;
        }
        if (endSymbols.contains(startSymbols.get(j))) {
          return (i == 0
              || Character.isWhitespace(txt.charAt(i - 1))
              || (i < txt.length() - 1 && !Character.isWhitespace(txt.charAt(i + 1))
                  && ((!PUNCT_MARKS.matcher(txt.substring(i + 1, i + 2)).matches()
                      && PUNCTUATION.matcher(txt.substring(i - 1, i)).matches())
                      || (txt.charAt(i + 1) == '-')))) ? j : -1;
        }
        return j;
      }
    }
    return -1;
  }

  /**
   * returns the index of the closing quote
   * -1 else
   */
  private int isClosingQuote(String txt, int i) throws Throwable {
    String tChar = txt.substring(i, i + 1);
    for (int j = 0; j < endSymbols.size(); j++) {
      if (endSymbols.get(j).equals(tChar)) {
        if (isNotQuote(txt, i, j) || nQuote != j) {
          return -1;
        }
        return j;
      }
    }
    return -1;
  }
  
  /**
   * True, if the quotation mark is a potential inch mark.
   */
  boolean isPotentialInchMark(String txt, int i) throws Throwable {
    if(!txt.substring(i, i + 1).equals("\"") || i < 1 || !Character.isDigit(txt.charAt(i - 1))) {
      return false;
    }
    int j;
    for (j = i - 1; j >= 0 && Character.isDigit(txt.charAt(j)); j--);
    if(j < 0 || Character.isWhitespace(txt.charAt(j))) {
      if(i + 1 < txt.length() && !txt.substring(i + 1, i + 2).equals(",")) {
        return true;
      }
    }
    return false;
  }

  /**
   * returns the index of the closing quote
   * -1 else
   */
  private int isPotentialClosingQuote(String txt, int i) throws Throwable {
    String tChar = txt.substring(i, i + 1);
    for (int j = 0; j < endSymbols.size(); j++) {
      if (endSymbols.get(j).equals(tChar)) {
        if (isNotQuote(txt, i, j) || isPotentialInchMark(txt, i)) {
          return -1;
        }
        return j;
      }
    }
    return -1;
  }
  
  /**
   * analyse one paragraph
   */
  private boolean analyzeOneParagraph(String txt, boolean isQuoteBefore) throws Throwable {
    openingQuotes = new ArrayList<>();
    closingQuotes = new ArrayList<>();
    if (isQuoteBefore) {
      openingQuotes.add(-1);
    }
    for (int i = 0; i < txt.length(); i++) {
      int tQuote = isOpeningQuote(txt, i);
      if (tQuote >= 0) {
        if (i == 0 && openingQuotes.size() > 0) {
          openingQuotes.set(0, i);
        } else {
          openingQuotes.add(i);
        }
        nQuote = tQuote;
        isQuoteBefore = true;
      } else {
        tQuote = isClosingQuote(txt, i);
        if (tQuote >= 0) {
          closingQuotes.add(i);
          nQuote = -1;
          isQuoteBefore = false;
        }
      }
    }
    return isQuoteBefore;
  }
  
  /**
   * analyse all paragraphs
   * NOTE: only text paragraphs will be analyzed (no footnotes, tables, etc.)
   * NOTE: Quotes lists have to initialized
   */
  public void analyzeTextParagraphs(List<String> paragraphs, List<List<Integer>> oQuotes, List<List<Integer>> cQuotes) {
    try {
      oQuotes.clear();
      cQuotes.clear();
      boolean isQuoteBefore = false;
      for (int i = 0; i < paragraphs.size(); i++) {
        isQuoteBefore = analyzeOneParagraph(paragraphs.get(i), isQuoteBefore);
        oQuotes.add(openingQuotes);
        cQuotes.add(closingQuotes);
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }

  /**
   * update a following paragraph
   * returns true, if the following paragraph has to be changed
   */
  private boolean updateFollowingParagraph(List<List<Integer>> oQuotes, List<List<Integer>> cQuotes, int nPara, boolean isQuoteBefore) {
    if (isQuoteBefore) {
      if (oQuotes.get(nPara) != null && oQuotes.get(nPara).size() > 0 && oQuotes.get(nPara).get(0) <= 0) {
        return false;
      }
      if (oQuotes.get(nPara) == null) {
        oQuotes.set(nPara, new ArrayList<Integer>());
      }
      oQuotes.get(nPara).add(0, -1);
      if (cQuotes.get(nPara) != null && cQuotes.get(nPara).size() > 0) {
        return false;
      } else {
        return true;
      }
    } else {
      if (oQuotes.get(nPara) == null || oQuotes.get(nPara).size() == 0 || oQuotes.get(nPara).get(0) >= 0) {
        return false;
      }
      oQuotes.get(nPara).remove(0);
      if (cQuotes.get(nPara) == null || cQuotes.get(nPara).size() == 0) {
        return true;
      } else {
        return false;
      }
    }
    
  }
  
  private void updateTextParagraphs(List<List<Integer>> oQuotes, List<List<Integer>> cQuotes, int nPara, boolean isQuoteBefore) {
    boolean hasChanged = true;
    for (int i = nPara + 1; hasChanged && i < oQuotes.size(); i++) {
      hasChanged = updateFollowingParagraph(oQuotes, cQuotes, i, isQuoteBefore);
    }
  }
  
  public void updateTextParagraph(String txt, int nPara, List<List<Integer>> oQuotes, List<List<Integer>> cQuotes) throws Throwable {
    if (nPara < 0 || nPara >= oQuotes.size()) {
      return;
    }
    boolean isQuoteBefore = oQuotes.get(nPara) != null && oQuotes.get(nPara).size() > 0 && oQuotes.get(nPara).get(0) < 0; 
    isQuoteBefore = analyzeOneParagraph(txt, isQuoteBefore);
    oQuotes.set(nPara, openingQuotes);
    cQuotes.set(nPara, closingQuotes);
    updateTextParagraphs(oQuotes, cQuotes, nPara, isQuoteBefore);
  }
  
  /**
   * Function to detect wrong double quotes
   * @throws Throwable 
   */
  public int numNotCorrectQuotes(String txt, int nQuote) throws Throwable {
    this.nQuote = nQuote;
    int num = 0;
    for (int i = 0; i < txt.length(); i++) {
      int tQuote = isOpeningQuote(txt, i);
      if (tQuote >= 0) {
        if ("\"".equals(startSymbols.get(tQuote))) {
          num++;
        } else if (nQuote < 0) {
          nQuote = tQuote;
        } else if (nQuote != tQuote) {
          num++;
        }
      } else {
        tQuote = isPotentialClosingQuote(txt, i);
        if (tQuote >= 0) {
          if ("\"".equals(endSymbols.get(tQuote))) {
            num++;
          } else if (nQuote != tQuote) {
            num++;
          }
        }
      }
    }
    return num;
  }

  /**
   * Function change wrong double quotes to correct quotes
   * gives back the corrected String
   * @throws Throwable 
   */
  public String changeToCorrectQuote(String txt) throws Throwable {
    return changeToCorrectQuote(txt, 0);
  }
  
  public String changeToCorrectQuote(String txt, int iCorrectQuote) throws Throwable {
    nChanges = 0;
    if (iCorrectQuote < 0 || iCorrectQuote >= startSymbols.size()) {
      return txt;
    }
    String out = new String(txt);
    for (int i = 0; i < txt.length(); i++) {
      int tQuote = isOpeningQuote(txt, i);
      if (tQuote >= 0) {
        if (tQuote != iCorrectQuote) {
          out = out.substring(0, i) + startSymbols.get(iCorrectQuote);
          if (i < txt.length() - 1) {
            out += txt.substring(i + 1);
          }
          nChanges++;
        }
      } else {
        tQuote = isPotentialClosingQuote(txt, i);
        if (tQuote >= 0) {
          if (tQuote != iCorrectQuote) {
            out = out.substring(0, i) + endSymbols.get(iCorrectQuote);
            if (i < txt.length() - 1) {
              out += txt.substring(i + 1);
            }
            nChanges++;
          }
        }
      }
    }
    return out;
  }

  /**
   * get the first possible correct Quote
   * @throws Throwable 
   */
  public int getFirstQuote(String txt) throws Throwable {
    for (int i = 0; i < txt.length(); i++) {
      int tQuote = isOpeningQuote(txt, i);
      if (tQuote >= 0) {
        if (tQuote < startSymbols.size() - 1) {
          return tQuote;
        }
      } else {
        tQuote = isPotentialClosingQuote(txt, i);
        if (tQuote >= 0) {
          if (tQuote < startSymbols.size() - 1) {
            return tQuote;
          }
        }
      }
    }
    return -1;
  }

  /**
   * get the detected possible correct Quote
   */
  public int getDetectedQuote() {
    return nQuote;
  }

  /**
   * get the number of changes
   */
  public int getNumChanges() {
    return nChanges;
  }

  private void setConfiguredQuotes(XComponentContext xContext, Locale locale) throws Throwable {

    XMultiServiceFactory xMSF = UnoRuntime.queryInterface(XMultiServiceFactory.class, xContext.getServiceManager());
    if (xMSF == null) {
      return;
    }
    Object oConfigProvider = xMSF.createInstance("com.sun.star.configuration.ConfigurationProvider");
    if (oConfigProvider == null) {
      WtMessageHandler.printToLogFile("oConfigProvider == null");
      return;
    }
    XMultiServiceFactory confMsf = UnoRuntime.queryInterface(XMultiServiceFactory.class, oConfigProvider);
        
    PropertyValue[] args = new PropertyValue[1];
    args[0] = new com.sun.star.beans.PropertyValue();
    args[0].Name = "nodepath";
    args[0].Value = "/org.openoffice.Office.Common/AutoCorrect";
    
    Object access = confMsf.createInstanceWithArguments(
        "com.sun.star.configuration.ConfigurationAccess", args);

    XPropertySet props = UnoRuntime.queryInterface(XPropertySet.class, access);
    
    int iDoubleStart = (int) props.getPropertyValue("DoubleQuoteAtStart");
    int iDoubleEnd   = (int) props.getPropertyValue("DoubleQuoteAtEnd");
//    int iSingleStart = (int) props.getPropertyValue("SingleQuoteAtStart");
//    int iSingleEnd   = (int) props.getPropertyValue("SingleQuoteAtEnd");
    
    String doubleStart = "";
    String doubleEnd = "";

    if (iDoubleStart == 0 || iDoubleEnd == 0) {
      Object olocaleData = xMSF.createInstance("com.sun.star.i18n.LocaleData");
      if (olocaleData == null) {
        WtMessageHandler.printToLogFile("olocaleData == null");
        return;
      }
      XLocaleData xlocaleData = UnoRuntime.queryInterface(XLocaleData.class, olocaleData);
      if (locale == null) {
        locale = WtOfficeTools.getDefaultLocale(xContext);
      }
      LocaleDataItem localeDataItem = xlocaleData.getLocaleItem(locale);
      doubleStart = localeDataItem.doubleQuotationStart;
      doubleEnd = localeDataItem.doubleQuotationEnd;
//      singleStart = localeDataItem.quotationStart.charAt(0);
//      singleEnd = localeDataItem.quotationEnd.charAt(0);
    }
    if (iDoubleStart != 0) {
      doubleStart = "" + ((char) iDoubleStart);
    }
    if (iDoubleEnd != 0) {
      doubleEnd = "" + ((char) iDoubleEnd);
    }
//    WtMessageHandler.printToLogFile("WtQuotesDetection: readQuotes: doubleStart: " + doubleStart + ", doubleEnd: " + doubleEnd);
//        + ", singleStart: " + singleStart + ", singleEnd: " + singleEnd);
    for (int i = startSymbols.size() - 1; i >= 0; i--) {
      if (startSymbols.get(i).equals(doubleStart)) {
        if (endSymbols.get(i).equals(doubleEnd)) {
          startSymbols.remove(i);
          endSymbols.remove(i);
        }
        break;
      }
    }
    startSymbols.add(0, doubleStart);
    endSymbols.add(0, doubleEnd);
    nQuote = 0;
  }

}
