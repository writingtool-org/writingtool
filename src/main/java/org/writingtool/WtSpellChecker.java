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

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.languagetool.AnalyzedSentence;
import org.languagetool.AnalyzedToken;
import org.languagetool.AnalyzedTokenReadings;
import org.languagetool.JLanguageTool;
import org.languagetool.JLanguageTool.ParagraphHandling;
import org.languagetool.Language;
import org.languagetool.rules.Rule;
import org.languagetool.rules.RuleMatch;
import org.languagetool.rules.spelling.SpellingCheckRule;
import org.languagetool.rules.spelling.hunspell.HunspellRule;
import org.languagetool.rules.spelling.morfologik.MorfologikSpellerRule;
import org.writingtool.WtCacheIO.SpellCache;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtVersionInfo;

import com.sun.star.beans.PropertyValue;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.Locale;
import com.sun.star.lang.XServiceDisplayName;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.lib.uno.helper.WeakBase;
import com.sun.star.linguistic2.XSpellAlternatives;
import com.sun.star.linguistic2.XSpellChecker;
import com.sun.star.uno.XComponentContext;

/**
 * Class for spell checking by LanguageTool
 * @since 1.0
 * @author Fred Kruse
 */
public class WtSpellChecker extends WeakBase implements XServiceInfo, 
  XServiceDisplayName, XSpellChecker {

  private static boolean DEBUG_MODE = true;  // set to true for debug output
  
  private static final int MAX_WRONG = 10000;
  private static final String PROB_CHARS = ".*[~<>].*";
  
  // Service name required by the OOo API && our own name.
  private static final String[] SERVICE_NAMES = {
          "com.sun.star.linguistic2.SpellChecker",
          WtOfficeTools.WT_SPELL_SERVICE_NAME };
  
  private static JLanguageTool lt = null;
  private static Locale lastLocale = null;                //  locale for spell check
  private Language lang;
  private static Language fixedLanguage;
  private static SpellingCheckRule spellingCheckRule = null;
  private static MorfologikSpellerRule mSpellRule = null;
  private static HunspellRule hSpellRule = null;
  private static final Map<String, List<String>> lastWrongWords = new HashMap<>();
  private static final Map<String, List<String[]>> lastSuggestions = new HashMap<>();
  private static String last1 = new String();
  private static String last2 = new String();
  private static XComponentContext xContext = null;
  private static boolean noLtSpeller = false;
  
  public WtSpellChecker(XComponentContext xContxt) {
    if (xContext == null) {
      try {
        xContext = xContxt;
        WtOfficeTools.renameOldLtFiles();     // This has to be deleted in later versions 
        WtMessageHandler.init(xContext, true);
        WtVersionInfo.init(xContext);
        if (WtVersionInfo.osArch == null || WtVersionInfo.osArch.equals("x86")
            || !isEnoughHeap() || !runLTSpellChecker(xContext)) {
          noLtSpeller = true;
        } else {
          WtCacheIO c = new WtCacheIO(); 
          SpellCache sc = c.new SpellCache();
          if (sc.read()) {
            if (sc.getWrongWords() != null && sc.getSuggestions() != null
                && sc.getWrongWords().size() == sc.getSuggestions().size()) {
              for (String loc : sc.getWrongWords().keySet()) {
                if (!sc.getWrongWords().get(loc).isEmpty()) {
                  List<String> savedLastWords = sc.getWrongWords().get(loc);
                  List<String[]> savedSuggestions = sc.getSuggestions().get(loc);
                  List<String> lastWords = new ArrayList<>();
                  List<String[]> suggestions = new ArrayList<>();
                  for (int i = 0; i < sc.getWrongWords().get(loc).size() && i < MAX_WRONG; i++) {
                    lastWords.add(savedLastWords.get(i));
                    suggestions.add(savedSuggestions.get(i));
                  }
                  lastWrongWords.put(loc, lastWords);
                  lastSuggestions.put(loc, suggestions);
                }
              }
            }
          }
        }
      } catch (Throwable e) {
        WtMessageHandler.showSpellError(e);
      }
    }
  }

  @Override
  public Locale[] getLocales() {
    if (noLtSpeller) {
      return new Locale[0];
    }
    return WtDocumentsHandler.getLocales();
  }

  @Override
  public boolean hasLocale(Locale locale) {
    return WtDocumentsHandler.hasLocale(locale);
  }

  @Override
  public String getImplementationName() {
    return WtSpellChecker.class.getName();
  }

  /**
   * Get the names of supported services
   * interface: XServiceInfo
   */
  @Override
  public String[] getSupportedServiceNames() {
    return getServiceNames();
  }

  /**
   * Get the LT service names
   */
  static String[] getServiceNames() {
    return SERVICE_NAMES;
  }

  /**
   * Test if the service is supported by LT
   * interface: XServiceInfo
   */
  @Override
  public boolean supportsService(String sServiceName) {
    for (String sName : SERVICE_NAMES) {
      if (sServiceName.equals(sName)) {
        return true;
      }
    }
    return false;
  }

  /**
   * is a correct spelled word
   */
  @Override
  public boolean isValid(String word, Locale locale, PropertyValue[] Properties) throws IllegalArgumentException {
    if (DEBUG_MODE) {
      WtMessageHandler.printToSpellLogFile("LtSpellChecker: isValid: test word: " + (word == null ? "null" : word)
          + ", DEBUG_MODE: " + DEBUG_MODE);
    }
    try {
      if (fixedLanguage != null) {
        locale = WtOfficeTools.getLocalFromLanguage(fixedLanguage);
      }
      if (noLtSpeller || locale == null || !hasLocale(locale)) {
        if (DEBUG_MODE) {
          WtMessageHandler.printToSpellLogFile("LtSpellChecker: isValid: noLtSpeller || locale == null || !hasLocale(locale): return false");
        }
        return false;
      }
      if (word == null || word.trim().isEmpty()) {
        if (DEBUG_MODE) {
          WtMessageHandler.printToSpellLogFile("LtSpellChecker: isValid: word id empty: " + (word == null ? "null" : word));
        }
        return true;
      }
      if (DEBUG_MODE) {
        WtMessageHandler.printToSpellLogFile("LtSpellChecker: isValid: test word/string: '" + (word == null ? "null" : word) + "'");
      }
      if (word.matches(PROB_CHARS)) {
        if (DEBUG_MODE) {
          WtMessageHandler.printToSpellLogFile("LtSpellChecker: isValid: Problematic word found: " + (word == null ? "null" : word));
        }
        return false;
      }
      String[] words = word.split(" ");
      if (words.length == 2 && words[0].equals(words[1])) {  //  this a workaround for a possible problem in LO
        if (last2.endsWith(".") || words[0].equals(last1)) {
          word = words[0] + " " + last2;
          if (DEBUG_MODE) {
            WtMessageHandler.printToSpellLogFile("LtSpellChecker: isValid: repeated word changed to: " + (word == null ? "null" : word));
          }
        } else {
          if (DEBUG_MODE) {
            WtMessageHandler.printToSpellLogFile("LtSpellChecker: isValid: repeated word found: " + (word == null ? "null" : word));
          }
          return false;
        }
      }
      last1 = last2;
      last2 = word;
      String localeStr = WtOfficeTools.localeToString(locale);
      List<String> wrongWords = lastWrongWords.get(localeStr);
      if (wrongWords != null && wrongWords.contains(word)) {
        if (DEBUG_MODE) {
          WtMessageHandler.printToSpellLogFile("LtSpellChecker: isValid: invalid word found in list: " + (word == null ? "null" : word));
        }
        return false;
      }
      initSpellChecker(locale);
      if (DEBUG_MODE) {
        WtMessageHandler.printToSpellLogFile("LtSpellChecker: isValid: spellingCheckRule is " + (spellingCheckRule == null ? "null" : "NOT null"));
      }
      if (spellingCheckRule != null) {
        if (words.length == 1 && !spellingCheckRule.isMisspelled(word)) {
          if (DEBUG_MODE) {
            WtMessageHandler.printToSpellLogFile("LtSpellChecker: isValid: valid word found: " + (word == null ? "null" : word));
          }
          return true;
        }

        List<RuleMatch> matches = lt.check(word,true, ParagraphHandling.ONLYNONPARA);
//        RuleMatch matches[] = spellingCheckRule.match(getAnalyzedSentence(word));

//        if (matches == null || matches.length == 0) {
        if (matches == null || matches.isEmpty()) {
          if (DEBUG_MODE) {
            WtMessageHandler.printToSpellLogFile("LtSpellChecker: isValid: valid word found (matches == 0): " + (word == null ? "null" : word));
          }
          return true;
        }
        if (wrongWords == null) {
          lastWrongWords.put(localeStr, new ArrayList<String>());
          lastSuggestions.put(localeStr, new ArrayList<String[]>());
        }
        if (!lastWrongWords.get(localeStr).contains(word)) {
          lastWrongWords.get(localeStr).add(new String(word));
//          lastSuggestions.get(localeStr).add(suggestionsToArray(matches[0].getSuggestedReplacements()));
          lastSuggestions.get(localeStr).add(suggestionsToArray(matches.get(0).getSuggestedReplacements()));
          if (lastWrongWords.get(localeStr).size() >= MAX_WRONG) {
            lastWrongWords.get(localeStr).remove(0);
            lastSuggestions.get(localeStr).remove(0);
          }
        }
        if (DEBUG_MODE) {
          WtMessageHandler.printToSpellLogFile("LtSpellChecker: isValid: invalid word found: " + (word == null ? "null" : word));
        }
        return false;
      }
    } catch (Throwable e) {
      WtMessageHandler.showSpellError(e);
    }
    WtMessageHandler.printToSpellLogFile("LtSpellChecker: isValid: Problem with spelling rule: word: " + (word == null ? "null" : word));
    return false;
  }

  /**
   * get alternatives for a non correct spelled word
   */
  @Override
  public XSpellAlternatives spell(String word, Locale locale, PropertyValue[] properties) throws IllegalArgumentException {
    if (fixedLanguage != null) {
      locale = WtOfficeTools.getLocalFromLanguage(fixedLanguage);
    }
    LTSpellAlternatives alternatives = new LTSpellAlternatives(word, locale);
    return alternatives;
  }

  /**
   * Init spell checker for locale
   */
  private void initSpellChecker(Locale locale) {
    try {
      if (lastLocale == null || lang == null || !WtOfficeTools.isEqualLocale(lastLocale, locale)) {
        lastLocale = locale;
//        WtMessageHandler.printToSpellLogFile("LtSpellChecker: initSpellChecker: fixed language: " + (fixedLanguage == null ? "null" : fixedLanguage.getShortCodeWithCountryAndVariant()));
        if(fixedLanguage != null) {
          lang = fixedLanguage;
        } else {
          lang = WtDocumentsHandler.getLanguage(locale);
        }
        lt = new JLanguageTool(lang);
        for (Rule rule : lt.getAllRules()) {
          if (rule.isDictionaryBasedSpellingRule()) {
            spellingCheckRule = (SpellingCheckRule) rule;
            if (spellingCheckRule instanceof MorfologikSpellerRule) {
              mSpellRule = (MorfologikSpellerRule) spellingCheckRule;
              hSpellRule = null;
            } else if (spellingCheckRule instanceof HunspellRule) {
              hSpellRule = (HunspellRule) spellingCheckRule;
              mSpellRule = null;
            }
          } else {
            lt.disableRule(rule.getId());
          }
        }
/*
        spellingCheckRule = lang.getDefaultSpellingRule();
        if (spellingCheckRule instanceof MorfologikSpellerRule) {
          mSpellRule = (MorfologikSpellerRule) spellingCheckRule;
          hSpellRule = null;
        } else if (spellingCheckRule instanceof HunspellRule) {
          hSpellRule = (HunspellRule) spellingCheckRule;
          mSpellRule = null;
        }
*/        
      }
    } catch (Throwable e) {
      WtMessageHandler.showSpellError(e);
    }
  }
  
  /**
   * get an analysed sentences from string
   * @throws IOException 
   */
  public AnalyzedSentence getAnalyzedSentence(String sentence) throws IOException {
    List<String> tokens = lang.getWordTokenizer().tokenize(sentence);

    List<AnalyzedTokenReadings> aTokens = lang.getTagger().tag(tokens);

    if (DEBUG_MODE) {
      WtMessageHandler.printToSpellLogFile("LtSpellChecker: getAnalyzedSentence: sentence: " + sentence + ", num tokens: " + tokens.size());
    }
    AnalyzedTokenReadings[] tokenArray = new AnalyzedTokenReadings[tokens.size() + 1];
    AnalyzedToken[] startTokenArray = new AnalyzedToken[1];
    int toArrayCount = 0;
    AnalyzedToken sentenceStartToken = new AnalyzedToken("", JLanguageTool.SENTENCE_START_TAGNAME, null);
    startTokenArray[0] = sentenceStartToken;
    tokenArray[toArrayCount++] = new AnalyzedTokenReadings(startTokenArray, 0);
    int startPos = 0;
    for (AnalyzedTokenReadings posTag : aTokens) {
      posTag.setStartPos(startPos);
      tokenArray[toArrayCount++] = posTag;
      startPos += posTag.getToken().length();
    }

    int numTokens = aTokens.size();
    int posFix = 0;
    for (int i = 0; i < numTokens; i++) {
      if (i > 0) {
        aTokens.get(i).setWhitespaceBefore(aTokens.get(i - 1).getToken());
        aTokens.get(i).setStartPos(aTokens.get(i).getStartPos() + posFix);
        aTokens.get(i).setPosFix(posFix);
      }
    }
    if (DEBUG_MODE) {
      WtMessageHandler.printToSpellLogFile("LtSpellChecker: getAnalyzedSentence: tokenArray length: " + tokenArray.length);
    }
    return new AnalyzedSentence(tokenArray);
  }

  /**
   * Convert list of suggestions to array and reduce it size
   */
  private String[] suggestionsToArray(List<String> suggestions) {
    int numSuggestions = suggestions.size();
    String[] allSuggestions = suggestions.toArray(new String[numSuggestions]);
    if (allSuggestions.length > WtOfficeTools.MAX_SUGGESTIONS) {
      allSuggestions = Arrays.copyOfRange(allSuggestions, 0, WtOfficeTools.MAX_SUGGESTIONS);
    }
    return allSuggestions;
  }
  
  /**
   * class for getting spelling alternatives
   */
  class LTSpellAlternatives implements XSpellAlternatives {
    
    Locale locale;
    String word;
    String[] alternatives = null;
    
    LTSpellAlternatives(String word, Locale locale) {
      try {
        this.word = word;
        this.locale = locale;
        if (noLtSpeller) {
          alternatives = new String[0];
          return;
        }
        String localeStr = WtOfficeTools.localeToString(locale);
        if (word == null || word.trim().isEmpty() || localeStr == null || localeStr.isEmpty()) {
          alternatives = new String[0];
          return;
        }
        List<String> wrongWords = lastWrongWords.get(localeStr);
        List<String[]> suggestions = lastSuggestions.get(localeStr);
        if (wrongWords != null && suggestions != null) {
          if (wrongWords.contains(word)) {
            int num = wrongWords.indexOf(word);
            alternatives = suggestions.get(num);
            if (alternatives == null) {
              alternatives = new String[0];
            }
            return;
          }
        }
        if (word.matches(PROB_CHARS)) {
          if (DEBUG_MODE) {
            WtMessageHandler.printToSpellLogFile("LtSpellChecker: LTSpellAlternatives: Problematic word found: " + (word == null ? "null" : word));
          }
          alternatives = new String[0];
          return;
        }
        if (mSpellRule != null) {
          alternatives = mSpellRule.getSpellingSuggestions(word).toArray(new String[0]);
        }
        if (hSpellRule != null) {
          alternatives = hSpellRule.getSuggestions(word).toArray(new String[0]);
        }
      } catch (Throwable t) {
        WtMessageHandler.showSpellError(t);
      }
      if (alternatives == null) {
        alternatives = new String[0];
      }
    }

    @Override
    public String[] getAlternatives() {
      return alternatives;
    }

    @Override
    public short getAlternativesCount() {
      return (short) alternatives.length;
    }

    @Override
    public short getFailureType() {
      // ??? unclear
      return 0;
    }

    @Override
    public Locale getLocale() {
      if (locale == null) {
        return lastLocale;
      }
      return locale;
    }

    @Override
    public String getWord() {
      if (word == null) {
        return "";
      }
      return word;
    }
    
  }

  @Override
  public String getServiceDisplayName(Locale locale) {
    return WtDocumentsHandler.getServiceDisplayName(locale);
  }
  
  public static Map<String, List<String>> getWrongWords() {
    return lastWrongWords;
  }
   
  public static Map<String, List<String[]>> getSuggestions() {
    return lastSuggestions;
  }
  
  public static boolean runLTSpellChecker(XComponentContext xContext) {
    WtConfiguration confg;
    try {
      confg = new WtConfiguration(WtOfficeTools.getWtConfigDir(xContext), 
          WtOfficeTools.CONFIG_FILE, null, true);
      fixedLanguage = confg.getDefaultLanguage();
      WtOfficeTools.setLogLevel(confg.getlogLevel());
      DEBUG_MODE = WtOfficeTools.DEBUG_MODE_SP;
      return confg.useLtSpellChecker();
    } catch (IOException e) {
      WtMessageHandler.printToSpellLogFile("Can't read configuration: LT spell checker not used!");
    }
    return false;
  }

  public static boolean isEnoughHeap() {
    int maxHeapSpace = (int) (WtOfficeTools.getMaxHeapSpace()/1048576);
    boolean ret = maxHeapSpace >= WtOfficeTools.SPELL_CHECK_MIN_HEAP;
    if (!ret) {
      WtMessageHandler.printToSpellLogFile("Heap Space (" + maxHeapSpace + ") is too small: LT spell checker not used!\n"
          + "Set heap space greater than " +  WtOfficeTools.SPELL_CHECK_MIN_HEAP);
    }
    return ret;
  }

  public static void resetSpellCache() {
    lastWrongWords.clear();
    lastSuggestions.clear();
  }
   
}
