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

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.ResourceBundle;
import java.util.Set;

import org.jetbrains.annotations.NotNull;
import org.languagetool.AnalyzedSentence;
import org.languagetool.AnalyzedToken;
import org.languagetool.AnalyzedTokenReadings;
import org.languagetool.JLanguageTool;
import org.languagetool.Language;
import org.languagetool.MultiThreadedJLanguageTool;
import org.languagetool.ResultCache;
import org.languagetool.ToneTag;
import org.languagetool.UserConfig;
import org.languagetool.JLanguageTool.Mode;
import org.languagetool.JLanguageTool.ParagraphHandling;
import org.languagetool.markup.AnnotatedTextBuilder;
import org.languagetool.rules.CategoryId;
import org.languagetool.rules.Rule;
import org.languagetool.rules.RuleMatch;
import org.writingtool.WtDocumentCache.AnalysedText;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.config.WtConfiguration;
import org.writingtool.remote.WtRemoteLanguageTool;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtOfficeTools.RemoteCheck;

/**
 * Class to switch between running LanguageTool in multi or single thread mode
 * @since 1.0
 * @author Fred Kruse
 */
public class WtLanguageTool {
  
  private static final ResourceBundle MESSAGES = WtOfficeTools.getMessageBundle();

  private final MultiThreadedJLanguageToolLo mlt;
  private final WtRemoteLanguageTool rlt;
  private JLanguageToolLo lt;
  private final Language language;
  private final Language motherTongue;
  private final UserConfig userConfig;

  private WtSortedTextRules sortedTextRules;
  private boolean isMultiThread;
  private boolean isRemote;
  private boolean doReset;
  private WtConfiguration config;

  public WtLanguageTool(Language language, Language motherTongue, UserConfig userConfig, 
      WtConfiguration config, List<Rule> extraRemoteRules,
      boolean noLtSpeller, boolean checkImpressDocument, boolean testMode) throws Throwable {
    this (language, motherTongue, userConfig, config, extraRemoteRules,
      noLtSpeller, checkImpressDocument, testMode, false);
  }

  public WtLanguageTool(Language language, Language motherTongue, UserConfig userConfig, 
      WtConfiguration config, List<Rule> extraRemoteRules,
      boolean noLtSpeller, boolean checkImpressDocument, boolean testMode, boolean onlyDefault) throws Throwable {
    this.config = config;
    this.language = language;
    this.motherTongue = motherTongue;
    this.userConfig = userConfig;
    isMultiThread = config.isMultiThread();
    isRemote = config.doRemoteCheck() && !testMode;
    doReset = false;
    if (isRemote) {
      lt = new JLanguageToolLo(language, motherTongue, null, userConfig);  // is needed to initialize synonym check
      mlt = null;
      rlt = new WtRemoteLanguageTool(language, motherTongue, config, extraRemoteRules, userConfig);
      //  TODO: CleanOverlappingMatches
      if (!rlt.remoteRun()) {
        WtMessageHandler.showMessage(MESSAGES.getString("loRemoteSwitchToLocal"));
        isRemote = false;
        isMultiThread = false;
      } else {
        lt = null;
      }
    } else if (isMultiThread) {
      lt = null;
      mlt = new MultiThreadedJLanguageToolLo(language, motherTongue, userConfig);
//      if (!config.filterOverlappingMatches()) {
        mlt.setCleanOverlappingMatches(false);
//      }
      rlt = null;
    } else {
      lt = new JLanguageToolLo(language, motherTongue, null, userConfig);
//      if (!config.filterOverlappingMatches()) {
        lt.setCleanOverlappingMatches(false);
//      }
      mlt = null;
      rlt = null;
    }
    File ngramDirectory = config.getNgramDirectory();
    if (ngramDirectory != null) {
      File ngramLangDir = new File(config.getNgramDirectory(), language.getShortCode());
      if (ngramLangDir.exists()) {  // user might have ngram data only for some languages and that's okay
        activateLanguageModelRules(ngramDirectory);
      }
    }
    if (noLtSpeller) {  // if LT spell checker is not use disable the spelling rules
      List<Rule> allRules = checkImpressDocument ? getAllActiveRules() : getAllActiveOfficeRules();
      for (Rule rule : allRules) {
        if (rule.isDictionaryBasedSpellingRule()) {
          disableRule(rule.getId());
          if (rule.useInOffice()) {
            // set default off so it can be re-enabled by user configuration
            rule.setDefaultOff();
          }
        }
      }
    }
    if (!onlyDefault) {
      initCheck(checkImpressDocument);
    }
  }
  
  /**
   * Enable or disable rules as given by configuration file
   */
  private void initCheck(boolean checkImpressDocument) throws Throwable {
    if (config.enableTmpOffRules()) {
      //  enable TempOff rules if configured
      List<Rule> allRules = getAllRules();
      WtMessageHandler.printToLogFile("initCheck: enableTmpOffRules: true");
      for (Rule rule : allRules) {
        if (rule.isDefaultTempOff()) {
          WtMessageHandler.printToLogFile("initCheck: enableTmpOffRule: " + rule.getId());
          enableRule(rule.getId());
        }
      }
    }
    Set<String> disabledRuleIds = config.getDisabledRuleIds();
    if (disabledRuleIds != null) {
      // copy as the config thread may access this as well
      List<String> list = new ArrayList<>(disabledRuleIds);
      for (String id : list) {
        disableRule(id);
      }
    }
    Set<String> disabledCategories = config.getDisabledCategoryNames();
    if (disabledCategories != null) {
      // copy as the config thread may access this as well
      List<String> list = new ArrayList<>(disabledCategories);
      for (String categoryName : list) {
        disableCategory(new CategoryId(categoryName));
      }
    }
    Set<String> enabledRuleIds = config.getEnabledRuleIds();
    if (enabledRuleIds != null) {
      // copy as the config thread may access this as well
      List<String> list = new ArrayList<>(enabledRuleIds);
      for (String ruleName : list) {
//        MessageHandler.printToLogFile("Enable Rule: " + ruleName);
        enableRule(ruleName);
      }
    }
    String langCode = getLanguage().getShortCodeWithCountryAndVariant();
    Set<String> disabledLocaleRules = WtDocumentsHandler.getDisabledRules(langCode);
    if (disabledLocaleRules != null) {
      for (String id : disabledLocaleRules) {
//        MessageHandler.printToLogFile("Disable local Rule: " + id + ", Locale: " + lt.getLanguage().getShortCodeWithCountryAndVariant());
        disableRule(id);
      }
    }
    sortedTextRules = new WtSortedTextRules(this, config, WtDocumentsHandler.getDisabledRules(langCode), checkImpressDocument);
//    handleLtDictionary();
  }
  
  /**
   * reset sorted text level rules
   */
  public void resetSortedTextRules(boolean checkImpressDocument) throws Throwable {
    String langCode = getLanguage().getShortCodeWithCountryAndVariant();
    sortedTextRules = new WtSortedTextRules(this, config, WtDocumentsHandler.getDisabledRules(langCode), checkImpressDocument);
  }

  /**
   * Returns a list of different numbers of paragraphs to check for text level rules
   */
  public List<Integer> getNumMinToCheckParas() {
    if (sortedTextRules == null) {
      return null;
    }
    return sortedTextRules.minToCheckParagraph;
  }

  /**
   * Test if sorted rules for index exist
   */
  public boolean isSortedRuleForIndex(int index) {
    if (index < 0 || sortedTextRules == null
        || index >= sortedTextRules.textLevelRules.size() || sortedTextRules.textLevelRules.get(index).isEmpty()) {
      return false;
    }
    return true;
  }

  /**
   * activate all rules stored under a given index related to the list of getNumMinToCheckParas
   * deactivate all other text level rules
   */
  public void activateTextRulesByIndex(int index) {
    if (sortedTextRules != null) {
      if (index == WtOfficeTools.CACHE_AI) {
        sortedTextRules.reactivateTextRules(this);
      } else {
        sortedTextRules.activateTextRulesByIndex(index, this);
      }
    }
  }

  /**
   * reactivate all text level rules
   */
  public void reactivateTextRules() {
    if (sortedTextRules != null) {
      sortedTextRules.reactivateTextRules(this);
    }
  }
  
  /**
   * Return true if check is done by a remote server
   */
  public boolean isRemote() {
    return isRemote;
  }

  /**
   * Get all rules
   */
  public List<Rule> getAllRules() {
    if (isRemote) {
      return rlt.getAllRules();
    } else if (isMultiThread) {
      return mlt.getAllRules(); 
    } else {
      return lt.getAllRules(); 
    }
  }

  /**
   * Get all active rules
   */
  public List<Rule> getAllActiveRules() {
    if (isRemote) {
      return rlt.getAllActiveRules();
    } else if (isMultiThread) {
        return mlt.getAllActiveRules(); 
    } else {
      return lt.getAllActiveRules(); 
    }
  }

  /**
   * Get all active office rules
   */
  public List<Rule> getAllActiveOfficeRules() {
    if (isRemote) {
      return rlt.getAllActiveOfficeRules();
    } else if (isMultiThread) {
        return mlt.getAllActiveOfficeRules(); 
    } else {
      return lt.getAllActiveOfficeRules(); 
    }
  }

  /**
   * Enable a rule by ID
   */
  public void enableRule(String ruleId) {
    if (isRemote) {
      rlt.enableRule(ruleId);
    } else if (isMultiThread) {
      mlt.enableRule(ruleId); 
    } else {
      lt.enableRule(ruleId); 
    }
  }

  /**
   * Disable a rule by ID
   */
  public void disableRule(String ruleId) {
    if (isRemote) {
      rlt.disableRule(ruleId);
    } else if (isMultiThread) {
      mlt.disableRule(ruleId); 
    } else {
      lt.disableRule(ruleId); 
    }
  }

  /**
   * Get disabled rules
   */
  public Set<String> getDisabledRules() {
    if (isRemote) {
      return rlt.getDisabledRules();
    } else if (isMultiThread) {
        return mlt.getDisabledRules(); 
    } else {
      return lt.getDisabledRules(); 
    }
  }

  /**
   * Disable a category by ID
   */
  public void disableCategory(CategoryId id) {
    if (isRemote) {
      rlt.disableCategory(id);
    } else if (isMultiThread) {
        mlt.disableCategory(id); 
    } else {
      lt.disableCategory(id); 
    }
  }

  /**
   * Activate language model (ngram) rules
   */
  public void activateLanguageModelRules(File indexDir) throws IOException {
    if (!isRemote) {
      if (isMultiThread) {
        mlt.activateLanguageModelRules(indexDir); 
      } else {
        lt.activateLanguageModelRules(indexDir); 
      }
    }
  }

  /**
   * check text by LT
   * default: check only grammar
   * local: LT checks only grammar (spell check is not implemented locally)
   * remote: spell checking is used for LT check dialog (is needed because method getAnalyzedSentence is not supported by remote check)
   */
  public List<RuleMatch> check(TextParagraph from, TextParagraph to, String text, ParagraphHandling paraMode, 
      WtSingleDocument document) throws IOException {
    return check(from, to, text, paraMode, document, RemoteCheck.ALL);
  }

  public List<RuleMatch> check(TextParagraph from, TextParagraph to, String text, ParagraphHandling paraMode, 
      WtSingleDocument document, RemoteCheck checkMode) throws IOException {
    if (isRemote) {
      List<RuleMatch> ruleMatches = rlt.check(text, paraMode, checkMode);
      if (ruleMatches == null) {
        doReset = true;
        ruleMatches = new ArrayList<>();
      }
      return ruleMatches;
    } else {
      Mode mode;
      if (paraMode == ParagraphHandling.ONLYNONPARA) {
        mode = Mode.ALL_BUT_TEXTLEVEL_ONLY;
      } else if (paraMode == ParagraphHandling.ONLYPARA) {
        mode = Mode.TEXTLEVEL_ONLY;
      } else {
        mode = Mode.ALL;
      }
      Set<ToneTag> toneTags = config.enableGoalSpecificRules() ? Collections.singleton(ToneTag.ALL_TONE_RULES) : Collections.emptySet();
      if (isMultiThread) {
        synchronized(mlt) {
          return mlt.check(from, to, paraMode, mode, document, this, toneTags);
        }
      } else {
        return lt.check(from, to, paraMode, mode, document, this, toneTags);
      }
    }
  }

  public List<RuleMatch> check(String text, ParagraphHandling paraMode, 
      int nFPara, WtSingleDocument document) throws IOException {
    return check(text, paraMode, nFPara, document, RemoteCheck.ALL);
  }

  public List<RuleMatch> check(String text, ParagraphHandling paraMode, 
      int nFPara, WtSingleDocument document, RemoteCheck checkMode) throws IOException {
    if (isRemote) {
      List<RuleMatch> ruleMatches = rlt.check(text, paraMode, checkMode);
      if (ruleMatches == null) {
        doReset = true;
        ruleMatches = new ArrayList<>();
      }
      return ruleMatches;
    } else {
      Mode mode;
      if (paraMode == ParagraphHandling.ONLYNONPARA) {
        mode = Mode.ALL_BUT_TEXTLEVEL_ONLY;
      } else if (paraMode == ParagraphHandling.ONLYPARA) {
        mode = Mode.TEXTLEVEL_ONLY;
      } else {
        mode = Mode.ALL;
      }
      Set<ToneTag> toneTags = config.enableGoalSpecificRules() ? Collections.singleton(ToneTag.ALL_TONE_RULES) : Collections.emptySet();
      if (isMultiThread) {
        synchronized(mlt) {
          return mlt.check(text, paraMode, mode, nFPara, document, this, toneTags);
        }
      } else {
        return lt.check(text, paraMode, mode, nFPara, document, this, toneTags);
      }
    }
  }

  public List<RuleMatch> check(String text, List<AnalyzedSentence> analyzedSentences, 
      ParagraphHandling paraMode, RemoteCheck checkMode) throws IOException {
    if (isRemote) {
      List<RuleMatch> ruleMatches = rlt.check(text, paraMode, checkMode);
      if (ruleMatches == null) {
        doReset = true;
        ruleMatches = new ArrayList<>();
      }
      return ruleMatches;
    } else {
      Mode mode;
      if (paraMode == ParagraphHandling.ONLYNONPARA) {
        mode = Mode.ALL_BUT_TEXTLEVEL_ONLY;
      } else if (paraMode == ParagraphHandling.ONLYPARA) {
        mode = Mode.TEXTLEVEL_ONLY;
      } else {
        mode = Mode.ALL;
      }
      Set<ToneTag> toneTags = config.enableGoalSpecificRules() ? Collections.singleton(ToneTag.ALL_TONE_RULES) : Collections.emptySet();
      if (isMultiThread) {
        synchronized(mlt) {
          return mlt.check(text, analyzedSentences, paraMode, mode, toneTags);
        }
      } else {
        return lt.check(text, analyzedSentences, paraMode, mode, toneTags);
      }
    }
  }

  /**
   * Get a list of tokens from a sentence
   * This Method use local lt for remote checks
   */
  public List<String> sentenceTokenize(String text) {
    if (isRemote) {
      if (lt == null) {
        lt = new JLanguageToolLo(language, motherTongue, null, userConfig);
      }
      return lt.sentenceTokenize(text);
    } else if (isMultiThread) {
        return mlt.sentenceTokenize(text); 
    } else {
      return lt.sentenceTokenize(text); 
    }
  }

  /**
   * Analyze sentence
   * This Method use local lt for remote checks
   */
  public AnalyzedSentence getAnalyzedSentence(String sentence) throws IOException {
    if (isRemote) {
      if (lt == null) {
        lt = new JLanguageToolLo(language, motherTongue, null, userConfig);
      }
      return lt.getAnalyzedSentence(sentence);
    } else if (isMultiThread) {
        return mlt.getAnalyzedSentence(sentence); 
    } else {
      return lt.getAnalyzedSentence(sentence); 
    }
  }

  /**
   * Analyze text
   * This Method use local lt for remote checks
   */
  public List<AnalyzedSentence> analyzeText(String text) throws IOException {
    if (isRemote) {
      if (lt == null) {
        lt = new JLanguageToolLo(language, motherTongue, null, userConfig);
      }
      return lt.analyzeText(text);
    } else if (isMultiThread) {
        return mlt.analyzeText(text); 
    } else {
      return lt.analyzeText(text); 
    }
  }

  /**
   * get the lemmas of a word
   * @throws IOException 
   */
  public List<String> getLemmasOfWord(String word) throws IOException {
    List<String> lemmas = new ArrayList<String>();
    Language language = getLanguage();
    List<String> words = new ArrayList<>();
    words.add(word);
    List<AnalyzedTokenReadings> aTokens = language.getTagger().tag(words);
    for (AnalyzedTokenReadings aToken : aTokens) {
      List<AnalyzedToken> readings = aToken.getReadings();
      for (AnalyzedToken reading : readings) {
        String lemma = reading.getLemma();
        if (lemma != null) {
          lemmas.add(lemma);
        }
      }
    }
    return lemmas;
  }

  /**
   * get the lemmas of a word
   * @throws IOException 
   */
  public List<String> getLemmasOfParagraph(String para, int startPos) throws IOException {
    List<String> lemmas = new ArrayList<String>();
    List<AnalyzedSentence> sentences = analyzeText(para);
    int pos = 0;
    for (AnalyzedSentence sentence : sentences) {
      for (AnalyzedTokenReadings token : sentence.getTokens()) {
        if (pos + token.getStartPos() == startPos) {
          for (AnalyzedToken reading : token) {
            String lemma = reading.getLemma();
            if (lemma != null) {
              lemmas.add(lemma);
            }
          }
        }
      }
      pos += sentence.getCorrectedTextLength();
    }
    return lemmas;
  }

  /**
   * Get the language from LT
   */
  public Language getLanguage() {
    if (isRemote) {
      return rlt.getLanguage();
    } else if (isMultiThread) {
      return mlt.getLanguage(); 
    } else {
      return lt.getLanguage(); 
    }
  }
  
  /**
   * Set reset flag
   */
  public boolean doReset() {
    return doReset;
  }
  
  public class JLanguageToolLo extends JLanguageTool {

    public JLanguageToolLo(Language language, Language motherTongue, ResultCache cache, UserConfig userConfig) {
      super(language, motherTongue, cache, userConfig);
    }

    public List<RuleMatch> check(String text, ParagraphHandling paraMode, Mode mode, 
        int nFPara, WtSingleDocument document, WtLanguageTool lt, @NotNull Set<ToneTag> toneTags) throws IOException {

      List<AnalyzedSentence> analyzedSentences;
      List<String> sentences;
      if (nFPara < 0) {
        analyzedSentences = this.analyzeText(text);
        sentences = new ArrayList<>();
        for (AnalyzedSentence analyzedSentence : analyzedSentences) {
          sentences.add(analyzedSentence.getText());
        }
      } else {
        AnalysedText analysedText = document.getDocumentCache().getOrCreateAnalyzedParagraph(nFPara, lt);
        analyzedSentences = analysedText.analyzedSentences;
        sentences = analysedText.sentences;
        text = analysedText.text;
      }
      return checkInternal(new AnnotatedTextBuilder().addText(text).build(), paraMode, null, mode, 
          Level.PICKY, toneTags, null, sentences, analyzedSentences).getRuleMatches();
    }

    public List<RuleMatch> check(TextParagraph from, TextParagraph to, ParagraphHandling paraMode, Mode mode, 
        WtSingleDocument document, WtLanguageTool lt, @NotNull Set<ToneTag> toneTags) throws IOException {
      AnalysedText analysedText = document.getDocumentCache().getAnalyzedParagraphs(from, to, lt);
      if (analysedText == null) {
        return null;
      }
      List<AnalyzedSentence> analyzedSentences = analysedText.analyzedSentences;
      List<String> sentences = analysedText.sentences;
      String text = analysedText.text;
      return checkInternal(new AnnotatedTextBuilder().addText(text).build(), paraMode, null, mode, 
          Level.PICKY, toneTags, null, sentences, analyzedSentences).getRuleMatches();
    }

    public List<RuleMatch> check(String text, List<AnalyzedSentence> analyzedSentences, 
        ParagraphHandling paraMode, Mode mode, @NotNull Set<ToneTag> toneTags) throws IOException {

      List<String> sentences;
      sentences = new ArrayList<>();
      for (AnalyzedSentence analyzedSentence : analyzedSentences) {
        sentences.add(analyzedSentence.getText());
      }
      return checkInternal(new AnnotatedTextBuilder().addText(text).build(), paraMode, null, mode, 
          Level.PICKY, toneTags, null, sentences, analyzedSentences).getRuleMatches();
    }

  }

  public class MultiThreadedJLanguageToolLo extends MultiThreadedJLanguageTool {

    public MultiThreadedJLanguageToolLo(Language language, Language motherTongue, UserConfig userConfig) {
      super(language, motherTongue, userConfig);
    }

    public List<RuleMatch> check(String text, ParagraphHandling paraMode, Mode mode, 
        int nFPara, WtSingleDocument document, WtLanguageTool lt, @NotNull Set<ToneTag> toneTags) throws IOException {

      List<AnalyzedSentence> analyzedSentences;
      List<String> sentences;
      if (nFPara < 0) {
        analyzedSentences = this.analyzeText(text);
        sentences = new ArrayList<>();
        for (AnalyzedSentence analyzedSentence : analyzedSentences) {
          sentences.add(analyzedSentence.getText());
        }
      } else {
        AnalysedText analysedText = document.getDocumentCache().getOrCreateAnalyzedParagraph(nFPara, lt);
        analyzedSentences = analysedText.analyzedSentences;
        sentences = analysedText.sentences;
        text = analysedText.text;
      }
      return checkInternal(new AnnotatedTextBuilder().addText(text).build(), paraMode, null, mode, 
          Level.PICKY, toneTags, null, sentences, analyzedSentences).getRuleMatches();
    }

    public List<RuleMatch> check(TextParagraph from, TextParagraph to, ParagraphHandling paraMode, Mode mode, 
        WtSingleDocument document, WtLanguageTool lt, @NotNull Set<ToneTag> toneTags) throws IOException {
      AnalysedText analysedText = document.getDocumentCache().getAnalyzedParagraphs(from, to, lt);
      if (analysedText == null) {
        return null;
      }
      List<AnalyzedSentence> analyzedSentences = analysedText.analyzedSentences;
      List<String> sentences = analysedText.sentences;
      String text = analysedText.text;
      return checkInternal(new AnnotatedTextBuilder().addText(text).build(), paraMode, null, mode, 
          Level.PICKY, toneTags, null, sentences, analyzedSentences).getRuleMatches();
    }

    public List<RuleMatch> check(String text, List<AnalyzedSentence> analyzedSentences, 
        ParagraphHandling paraMode, Mode mode, @NotNull Set<ToneTag> toneTags) throws IOException {

      List<String> sentences;
      sentences = new ArrayList<>();
      for (AnalyzedSentence analyzedSentence : analyzedSentences) {
        sentences.add(analyzedSentence.getText());
      }
      return checkInternal(new AnnotatedTextBuilder().addText(text).build(), paraMode, null, mode, 
          Level.PICKY, toneTags, null, sentences, analyzedSentences).getRuleMatches();
    }

  }

}
