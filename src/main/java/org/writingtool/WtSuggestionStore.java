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
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

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

  private final Map<String, LinkedHashSet<String>> lastWords = new HashMap<>();
  private final Map<String, LinkedHashMap<String, String[]>> lastSuggestions = new HashMap<>();
  
  private final int maxWords;

  public WtSuggestionStore(int maxWords) {
    this.maxWords = maxWords;
  }
  
  /**
   * Is a word in store
   */
  public boolean hasWord(String word, Locale locale) throws Throwable {
    String localeStr = WtOfficeTools.localeToString(locale);
    Set<String> words = lastWords.get(localeStr);
    return words != null && words.contains(word);
  }
  
  /**
   * get the suggestions of a word as an array of strings
   */
  public String[] getSuggestions(String word, Locale locale) throws Throwable {
    String localeStr = WtOfficeTools.localeToString(locale);
    Set<String> words = lastWords.get(localeStr);
    if (words == null || !words.contains(word)) {
//      WtMessageHandler.printToLogFile("WtSuggestionStore: getSuggestions: " + (words == null ? "Words == null" : "word not found: " + word));
      return null;
    }
    LinkedHashMap<String, String[]> suggestionMap = lastSuggestions.get(localeStr);
    if (suggestionMap == null || suggestionMap.size() != words.size()) {
      WtMessageHandler.printToLogFile("WtSuggestionStore error: Size don't match! Clear store");
      lastWords.clear();
      lastSuggestions.clear();
      return null;
    }
    String[] suggestions = suggestionMap.remove(word);
    if (suggestions == null) {
      lastWords.clear();
      lastSuggestions.clear();
      return null;
    }
    words.remove(word);
    words.add(word);
    suggestionMap.put(word, suggestions);
    return suggestions;
  }
  
  /**
   * add the suggestions of word as array of strings
   */
  public void addSuggestions(String word, Locale locale, String[] suggestions) throws Throwable {
    String localeStr = WtOfficeTools.localeToString(locale);
    Set<String> words = lastWords.computeIfAbsent(localeStr, key -> new LinkedHashSet<>());
    LinkedHashMap<String, String[]> suggestionMap = lastSuggestions.computeIfAbsent(localeStr, key -> new LinkedHashMap<>());
    if (words.contains(word)) {
      words.remove(word);
      suggestionMap.remove(word);
    }
    words.add(word);
    suggestionMap.put(word, suggestions);
    while (words.size() > maxWords) {
      String oldestWord = words.iterator().next();
      words.remove(oldestWord);
      suggestionMap.remove(oldestWord);
    }
  }
  
  /**
   * Set the full store of words of suggestions
   * The store is cleared before
   */
  public void setFullStore(Map<String, List<String>> words, Map<String, List<String[]>> suggestions) throws Throwable {
    if (words == null || suggestions == null || words.size() != suggestions.size()) {
      return;
    }
    lastWords.clear();
    lastSuggestions.clear();
    for (Map.Entry<String, List<String>> entry : words.entrySet()) {
      String loc = entry.getKey();
      List<String> savedLastWords = entry.getValue();
      if (savedLastWords != null && !savedLastWords.isEmpty()) {
        List<String[]> savedSuggestions = suggestions.get(loc);
        if (savedSuggestions == null) {
          continue;
        }
        LinkedHashSet<String> lastSavedWords = new LinkedHashSet<>();
        LinkedHashMap<String, String[]> lastSavedSuggestions = new LinkedHashMap<>();
        for (int i = 0; i < savedLastWords.size() && i < maxWords; i++) {
          String savedWord = savedLastWords.get(i);
          if (savedWord != null) {
            lastSavedWords.add(savedWord);
            lastSavedSuggestions.put(savedWord, savedSuggestions.get(i));
          }
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
    Map<String, List<String>> result = new HashMap<>();
    for (Map.Entry<String, LinkedHashSet<String>> entry : lastWords.entrySet()) {
      result.put(entry.getKey(), new ArrayList<>(entry.getValue()));
    }
    return result;
  }

  /**
   * get the map of last suggestions (depends on locale)
   */
  public Map<String, List<String[]>> getLastSuggestions() {
    Map<String, List<String[]>> result = new HashMap<>();
    for (Map.Entry<String, LinkedHashMap<String, String[]>> entry : lastSuggestions.entrySet()) {
      result.put(entry.getKey(), new ArrayList<>(entry.getValue().values()));
    }
    return result;
  }
  
  /**
   * clear the whole store
   */
  public void clear() {
    lastWords.clear();
    lastSuggestions.clear();
  }
  
  

}
