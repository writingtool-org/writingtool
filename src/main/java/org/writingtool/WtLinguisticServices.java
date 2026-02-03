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
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.languagetool.Language;
import org.languagetool.LinguServices;
import org.languagetool.rules.Rule;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtLinguServiceTools;

import com.sun.star.beans.Property;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.lang.Locale;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.linguistic2.XHyphenator;
import com.sun.star.linguistic2.XLinguServiceManager;
import com.sun.star.linguistic2.XMeaning;
import com.sun.star.linguistic2.XPossibleHyphens;
import com.sun.star.linguistic2.XSpellAlternatives;
import com.sun.star.linguistic2.XSpellChecker;
import com.sun.star.linguistic2.XThesaurus;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
 * Class to handle information from linguistic services of LibreOffice/OpenOffice
 * @since 1.0
 * @author Fred Kruse
 */
public class WtLinguisticServices extends LinguServices {
  
//  private XThesaurus thesaurus = null;
//  private XSpellChecker spellChecker = null;
//  private XHyphenator hyphenator = null;
  private XComponentContext xContext;
  private Map<String, List<String>> synonymsCache;
  private List<String> thesaurusRelevantRules = null;
  private boolean noSynonymsAsSuggestions = false;

  public WtLinguisticServices(XComponentContext xContext) {
    this.xContext = xContext;
    synonymsCache = new HashMap<>();
//    if (xContext != null) {
//      XLinguServiceManager mxLinguSvcMgr = getLinguSvcMgr(xContext);
//      thesaurus = getThesaurus(mxLinguSvcMgr);
//      spellChecker = getSpellChecker(mxLinguSvcMgr);
//      hyphenator = getHyphenator(mxLinguSvcMgr);
//      synonymsCache = new HashMap<>();
//    }
  }

  /**
   * Set Parameter to generate no synonyms (makes some rules faster, but results have no suggestions)
   */
  public void setNoSynonymsAsSuggestions (boolean noSynonymsAsSuggestions) {
    this.noSynonymsAsSuggestions = noSynonymsAsSuggestions;
  }
  
  /**
   * returns if spell checker can be used
   * if false initialize LinguisticServices again
   *//*
  public boolean spellCheckerIsActive () {
    return (spellChecker != null);
  }
  */

  /** 
   * Get XLinguProperties
   */
  private static XPropertySet getLinguProperties(XComponentContext xContext) {
    if (xContext == null) {
      return null;
    }
    XMultiComponentFactory xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class, xContext.getServiceManager());
    if (xMCF == null) {
      return null;
    }
    Object linguProperties = null;
    try {
      linguProperties = xMCF.createInstanceWithContext("com.sun.star.linguistic2.LinguProperties", xContext);
    } catch (Exception e) {
      WtMessageHandler.printException(e);
    }
    if (linguProperties == null) {
      return null;
    }
    return UnoRuntime.queryInterface(XPropertySet.class, linguProperties);
  }
  
  /**
   * Print LiguProperties to log file (Used for tests only)
   */
  
  public void printLinguProperties(XComponentContext xContext) {
    XPropertySet propSet = getLinguProperties(xContext);
    XPropertySetInfo propertySetInfo = propSet.getPropertySetInfo();
    WtMessageHandler.printToLogFile("OfficeTools: printPropertySet: PropertySet:");
    for (Property property : propertySetInfo.getProperties()) {
      WtMessageHandler.printToLogFile("Name: " + property.Name + ", Type: " + property.Type.getTypeName());
    }
  }
  
  /** 
   * Get the Thesaurus to be used.
   */
  private XThesaurus getThesaurus(XComponentContext xContext) {
    try {
      XLinguServiceManager mxLinguSvcMgr = WtLinguServiceTools.getLinguSvcMgr(xContext);
      if (mxLinguSvcMgr != null) {
        return mxLinguSvcMgr.getThesaurus();
      }
    } catch (Throwable t) {
      // If anything goes wrong, give the user a stack trace
      WtMessageHandler.printException(t);
    }
    return null;
  }

  /** 
   * Get the Hyphenator to be used.
   */
  private XHyphenator getHyphenator(XComponentContext xContext) {
    try {
      XLinguServiceManager mxLinguSvcMgr = WtLinguServiceTools.getLinguSvcMgr(xContext);
      if (mxLinguSvcMgr != null) {
        return mxLinguSvcMgr.getHyphenator();
      }
    } catch (Throwable t) {
      // If anything goes wrong, give the user a stack trace
      WtMessageHandler.printException(t);
    }
    return null;
  }

  /** 
   * Get the SpellChecker to be used.
   */
  protected XSpellChecker getSpellChecker(XComponentContext xContext) {
    try {
      XLinguServiceManager mxLinguSvcMgr = WtLinguServiceTools.getLinguSvcMgr(xContext);
      if (mxLinguSvcMgr != null) {
        return mxLinguSvcMgr.getSpellChecker();
      }
    } catch (Throwable t) {
      // If anything goes wrong, give the user a stack trace
      WtMessageHandler.printException(t);
    }
    return null;
  }

  /**
   * Get a Locale from a LT defined language
   */
  public static Locale getLocale(Language lang) {
    if (lang == null) {
      return null;
    }
    Locale locale = new Locale();
    locale.Language = lang.getShortCode();
    if ((lang.getCountries() == null || lang.getCountries().length != 1) && lang.getDefaultLanguageVariant() != null) {
      locale.Country = lang.getDefaultLanguageVariant().getCountries()[0];
    } else if (lang.getCountries() != null && lang.getCountries().length > 0) {
      locale.Country = lang.getCountries()[0];
    } else {
      locale.Country = "";
    }
    if (lang.getVariant() == null) {
      locale.Variant = "";
    } else {
      locale.Variant = lang.getVariant();
    }
    return locale;
  }
  
  /**
   * Get all synonyms of a word as list of strings.
   */
  @Override
  public List<String> getSynonyms(String word, Language lang) {
    return getSynonyms(word, getLocale(lang));
  }
  
  public List<String> getSynonyms(String word, Locale locale) {
    try {
      if (noSynonymsAsSuggestions) {
        return new ArrayList<>();
      }
      if (synonymsCache.containsKey(word)) {
        return synonymsCache.get(word);
      }
      // get synonyms in a acceptable time or return 0 synonyms
      AddSynonymsToCache addSynonymsToCache = new AddSynonymsToCache(word, locale);
      addSynonymsToCache.start();
      long startTime = System.currentTimeMillis();
      long runTime = 0;
      do {
        Thread.sleep(10);
        if (synonymsCache.containsKey(word)) {
          return synonymsCache.get(word);
        }
        runTime = System.currentTimeMillis() - startTime;
      } while (runTime < 500);
    } catch (InterruptedException e) {
      WtMessageHandler.printException(e);
    }
    return new ArrayList<>();
  }
  
  /**
   * Returns true if the spell check is positive
   */
  @Override
  public boolean isCorrectSpell(String word, Language lang) {
    return isCorrectSpell(word, getLocale(lang));
  }
  
  public boolean isCorrectSpell(String word, Locale locale) {
    XSpellChecker spellChecker = getSpellChecker(xContext);
    if (spellChecker == null) {
      WtMessageHandler.printToLogFile("LinguisticServices: isCorrectSpell: XSpellChecker == null");
      return false;
    }
    PropertyValue[] properties = new PropertyValue[0];
    try {
      return spellChecker.isValid(word, locale, properties);
    } catch (Throwable t) {
      // If anything goes wrong, give the user a stack trace
      WtMessageHandler.printException(t);
      return false;
    }
  }

  /**
   * Returns Alternatives to  wrong spelled word
   */
  public String[] getSpellAlternatives(String word, Language lang) {
    return getSpellAlternatives(word, getLocale(lang));
  }
  
  public String[] getSpellAlternatives(String word, Locale locale) {
    XSpellChecker spellChecker = getSpellChecker(xContext);
    if (spellChecker == null) {
      WtMessageHandler.printToLogFile("LinguisticServices: getSpellAlternatives: XSpellChecker == null");
      return null;
    }
    PropertyValue[] properties = new PropertyValue[0];
    try {
      XSpellAlternatives spellAlternatives = spellChecker.spell(word, locale, properties);
      if (spellAlternatives == null) {
        return null;
      }
      return spellAlternatives.getAlternatives();
    } catch (Throwable t) {
      // If anything goes wrong, give the user a stack trace
      WtMessageHandler.printException(t);
      return null;
    }
  }

  /**
   * Returns the number of syllable of a word
   * Returns -1 if the word was not found or anything goes wrong
   */
  @Override
  public int getNumberOfSyllables(String word, Language lang) {
    return getNumberOfSyllables(word, getLocale(lang));
  }
  
  public int getNumberOfSyllables(String word, Locale locale) {
    XHyphenator hyphenator = getHyphenator(xContext);
    if (hyphenator == null) {
      WtMessageHandler.printToLogFile("LinguisticServices: getNumberOfSyllables: XHyphenator == null");
      return 1;
    }
    PropertyValue[] properties = new PropertyValue[0];
    try {
      XPossibleHyphens possibleHyphens = hyphenator.createPossibleHyphens(word, locale, properties);
      if (possibleHyphens == null) {
        return 1;
      }
      short[] numSyllable = possibleHyphens.getHyphenationPositions();
      return numSyllable.length + 1;
    } catch (Throwable t) {
      // If anything goes wrong, give the user a stack trace
      WtMessageHandler.printException(t);
      return 1;
    }
  }
  

  /**
   * Set a thesaurus relevant rule
   */
  @Override
  public void setThesaurusRelevantRule (Rule rule) {
    if (thesaurusRelevantRules == null) {
      thesaurusRelevantRules = new ArrayList<String>();
    }
    String ruleId = rule.getId();
    if (ruleId != null && !thesaurusRelevantRules.contains(ruleId)) {
      thesaurusRelevantRules.add(ruleId);
    }
  }

  /**
   * Test if rule is thesaurus relevant 
   * (match should give suggestions from thesaurus)
   */
  public boolean isThesaurusRelevantRule (String ruleId) {
    return !noSynonymsAsSuggestions && thesaurusRelevantRules != null && thesaurusRelevantRules.contains(ruleId);
  }
  
  /** class to start a separate thread to add Synonyms to cache
   *  To get synonyms in a acceptable time or return null
   */
  private class AddSynonymsToCache extends Thread {
    private String word;
    private Locale locale;

    private AddSynonymsToCache(String word, Locale locale) {
      this.word = word;
      this.locale = locale;
    }
    
    @Override
    public void run() {
      List<String> synonyms = new ArrayList<>();
      try {
        XThesaurus thesaurus = getThesaurus(xContext);
        if (thesaurus == null) {
          WtMessageHandler.printToLogFile("LinguisticServices: getSynonyms: XThesaurus == null");
          return;
        }
        if (locale == null) {
          WtMessageHandler.printToLogFile("LinguisticServices: getSynonyms: Locale == null");
          return;
        }
        PropertyValue[] properties = new PropertyValue[0];
        XMeaning[] meanings = thesaurus.queryMeanings(word, locale, properties);
        for (XMeaning meaning : meanings) {
          if (synonyms.size() >= WtOfficeTools.MAX_SUGGESTIONS) {
            break;
          }
          String[] singleSynonyms = meaning.querySynonyms();
          Collections.addAll(synonyms, singleSynonyms);
        }
        synonymsCache.put(word, synonyms);
      } catch (Throwable t) {
        // If anything goes wrong, give the user a stack trace
        WtMessageHandler.printException(t);
      }
    }

  }

}
