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
import java.util.List;

import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;

import com.sun.star.linguistic2.XDictionary;
import com.sun.star.linguistic2.XSearchableDictionaryList;
import com.sun.star.uno.XComponentContext;

/**
 * Class to add manual LT dictionaries temporarily to LibreOffice/OpenOffice
 * @since 1.0
 * @author Fred Kruse
 */
public class WtDictionary {
  
  private final static String INTERNAL_DICT_PREFIX = "__LT_";
  private final static String DICT_FILE_POSTFIX = ".dic";

  private static String listIgnoredWords = null;
  private static boolean activateDictionary = false;
  
  /**
   * Add a non permanent dictionary to LO/OO that contains additional words defined in LT
   */
  public static boolean isActivating() {
    return activateDictionary;
  }
  
  private static void setListIgnoredWords(XDictionary[] dictionaryList) throws Throwable {
    if (listIgnoredWords != null) {
      return;
    }
    boolean debugMode = WtOfficeTools.DEBUG_MODE_LD;
    for (XDictionary dictionary : dictionaryList) {
      if (dictionary.isActive()) {
        String name = dictionary.getName();
        if (!name.startsWith(INTERNAL_DICT_PREFIX) && !name.endsWith(DICT_FILE_POSTFIX)) {
          listIgnoredWords = new String(name);
          if (debugMode) {
            WtMessageHandler.printToLogFile("dictionary for ignored words found: " + listIgnoredWords);
          }
        }
      }
    }
    if (listIgnoredWords == null) {
      WtMessageHandler.printToLogFile("WARNING: dictionary for ignored words not found!");
    }
  }
  
  private static XDictionary getListIgnoredWords(XComponentContext xContext) throws Throwable {
    XSearchableDictionaryList searchableDictionaryList = WtOfficeTools.getSearchableDictionaryList(xContext);
    if (searchableDictionaryList == null) {
      WtMessageHandler.printToLogFile("LtDictionary: getListIgnoredWords: searchableDictionaryList == null");
      return null;
    }
    if (listIgnoredWords == null) {
      setListIgnoredWords(searchableDictionaryList.getDictionaries());
    }
    for (XDictionary dictionary : searchableDictionaryList.getDictionaries()) {
      if (dictionary.isActive() && listIgnoredWords.equals(dictionary.getName())) {
        return dictionary;
      }
    }
    WtMessageHandler.printToLogFile("WARNING: dictionary for ignored words not found!");
    return null;
  }

  /**
   * Add a word to the List of ignored words
   * Used for ignore all in spelling check
   */
  public static void addIgnoredWord(String word, XComponentContext xContext) throws Throwable {
    XDictionary ignoredWords = getListIgnoredWords(xContext);
    ignoredWords.add(word, false, "");
  }
  
  /**
   * Remove a word from the List of ignored words
   * Used for ignore all in spelling check
   */
  public static void removeIgnoredWord(String word, XComponentContext xContext) throws Throwable {
    XDictionary ignoredWords = getListIgnoredWords(xContext);
    ignoredWords.remove(word);
  }
  
  /**
   * Add a word to a user dictionary
   */
  public static void addWordToDictionary(String dictionaryName, String word, XComponentContext xContext) throws Throwable {
    if (word == null) {
      throw new RuntimeException("No word selected to add to dictionary (word == null) ");
    }
    if (dictionaryName == null) {
      throw new RuntimeException("No dictionary selected (dictionaryName == null) ");
    }
    XSearchableDictionaryList searchableDictionaryList = WtOfficeTools.getSearchableDictionaryList(xContext);
    if (searchableDictionaryList == null) {
      WtMessageHandler.printToLogFile("LtDictionary: addWordToDictionary: searchableDictionaryList == null");
      return;
    }
    XDictionary dictionary = searchableDictionaryList.getDictionaryByName(dictionaryName);
    if (dictionary == null) {
      throw new RuntimeException("Dictionary not found (dictionaryName: " + dictionaryName + "; word: " + word);
    }
    dictionary.add(word, false, "");
  }
  
  /**
   * Add a word to a user dictionary
   */
  public static void removeWordFromDictionary(String dictionaryName, String word, XComponentContext xContext) throws Throwable {
    XSearchableDictionaryList searchableDictionaryList = WtOfficeTools.getSearchableDictionaryList(xContext);
    if (searchableDictionaryList == null) {
      WtMessageHandler.printToLogFile("LtDictionary: removeWordFromDictionary: searchableDictionaryList == null");
      return;
    }
    XDictionary dictionary = searchableDictionaryList.getDictionaryByName(dictionaryName);
    dictionary.remove(word);
  }
  
  /**
   * Get all user dictionaries
   */
  public static String[] getUserDictionaries(XComponentContext xContext) throws Throwable {
    boolean debugMode = WtOfficeTools.DEBUG_MODE_LD;
    XSearchableDictionaryList searchableDictionaryList = WtOfficeTools.getSearchableDictionaryList(xContext);
    if (searchableDictionaryList == null) {
      WtMessageHandler.printToLogFile("LtDictionary: getUserDictionaries: searchableDictionaryList == null");
      return null;
    }
    XDictionary[] dictionaryList = searchableDictionaryList.getDictionaries();
    setListIgnoredWords(dictionaryList);
    List<String> userDictionaries = new ArrayList<String>();
    for (XDictionary dictionary : dictionaryList) {
      if (debugMode) {
        WtMessageHandler.printToLogFile("LtDictionary: getUserDictionaries: dictionary: " + dictionary.getName() 
            + ", isActive: " + dictionary.isActive() + ", Type: " + dictionary.getDictionaryType().getValue());
      }
      if (dictionary.isActive()) {
        String name = dictionary.getName();
        if (!name.startsWith(INTERNAL_DICT_PREFIX) && !name.equals(listIgnoredWords)) {
          userDictionaries.add(new String(name));
        }
      }
    }
    return userDictionaries.toArray(new String[userDictionaries.size()]);
  }
}
