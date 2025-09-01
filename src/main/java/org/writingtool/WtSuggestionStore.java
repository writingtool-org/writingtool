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
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;

import com.sun.star.lang.Locale;

/**
 * Class to store suggestions for a word 
 * depends on locale
 * stores maxWords words and its suggestions
 * if maxWords is exceeded the logest not used word and its suggestions are deleted
 * @since 10.7
 */
public class WtSuggestionStore {

  private final Map<String, List<String>> lastWords = new HashMap<>();
  private final Map<String, List<String[]>> lastSuggestions = new HashMap<>();
  
  private final int maxWords;

  public WtSuggestionStore(int maxWords) {
    this.maxWords = maxWords;
  }
  
  /**
   * Is a word in store
   */
  public boolean hasWord(String word, Locale locale) {
    String localeStr = WtOfficeTools.localeToString(locale);
    List<String> words = lastWords.get(localeStr);
    return words != null && words.contains(word);
  }
  
  /**
   * get the suggestions of a word as an array of strings
   */
  public String[] getSuggestions(String word, Locale locale) {
    String localeStr = WtOfficeTools.localeToString(locale);
    List<String> words = lastWords.get(localeStr);
    if (words == null || !words.contains(word)) {
//      WtMessageHandler.printToLogFile("WtSuggestionStore: getSuggestions: " + (words == null ? "Words == null" : "word not found: " + word));
      return null;
    }
    List<String[]> suggestionList = lastSuggestions.get(localeStr);
    if (suggestionList == null || suggestionList.size() != words.size()) {
      WtMessageHandler.printToLogFile("WtSuggestionStore error: Size don't match! Clear store");
      lastWords.clear();
      lastSuggestions.clear();
      return null;
    }
    int n = words.indexOf(word);
    String[] suggestions = suggestionList.get(n);
    words.remove(n);
    suggestionList.remove(n);
    words.add(word);
    suggestionList.add(suggestions);
    return suggestions;
  }
  
  /**
   * add the suggestions of word as array of strings
   */
  public void addSuggestions(String word, Locale locale, String[] suggestions) {
    String localeStr = WtOfficeTools.localeToString(locale);
    List<String> words = lastWords.get(localeStr);
    List<String[]> suggestionList;
    if (words == null) {
      words = new ArrayList<>();
      suggestionList = new ArrayList<>();
      lastWords.put(localeStr, words);
      lastSuggestions.put(localeStr, suggestionList);
    } else {
      suggestionList = lastSuggestions.get(localeStr);
    }
    if (words.contains(word)) {
      int n = words.indexOf(word);
      words.remove(n);
      suggestionList.remove(n);
    }
    words.add(word);
    suggestionList.add(suggestions);
    while (words.size() > maxWords) {
      words.remove(0);
      suggestionList.remove(0);
    }
    lastWords.put(localeStr, words);
    lastSuggestions.put(localeStr, suggestionList);
  }
  
  /**
   * Set the full store of words of suggestions
   * The store is cleared before
   */
  public void setFullStore(Map<String, List<String>> words, Map<String, List<String[]>> suggestions) {
    if (words == null || suggestions == null || words.size() != suggestions.size()) {
      return;
    }
    lastWords.clear();
    lastSuggestions.clear();
    for (String loc : words.keySet()) {
      if (!words.get(loc).isEmpty()) {
        List<String> savedLastWords = words.get(loc);
        List<String[]> savedSuggestions = suggestions.get(loc);
        List<String> lastSavedWords = new ArrayList<>();
        List<String[]> lastSavedSuggestions = new ArrayList<>();
        for (int i = 0; i < savedLastWords.size() && i < maxWords; i++) {
          lastSavedWords.add(savedLastWords.get(i));
          lastSavedSuggestions.add(savedSuggestions.get(i));
        }
        lastWords.put(loc, lastSavedWords);
        lastSuggestions.put(loc, lastSavedSuggestions);
      }
    }
  }
  
  /**
   * get the map of last words (depends on locale)
   */
  public Map<String, List<String>> getLastWords() {
    return lastWords;
  }

  /**
   * get the map of last suggestions (depends on locale)
   */
  public Map<String, List<String[]>> getLastSuggestions() {
    return lastSuggestions;
  }
  
  /**
   * clear the whole store
   */
  public void clear() {
    lastWords.clear();
    lastSuggestions.clear();
  }
  
  

}
