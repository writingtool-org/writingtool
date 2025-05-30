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
package org.writingtool.remote;

import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.languagetool.AnalyzedSentence;
// import org.languagetool.AnalyzedToken;
// import org.languagetool.AnalyzedTokenReadings;
// import org.languagetool.JLanguageTool;
import org.languagetool.Language;
// import org.languagetool.LinguServices;
import org.languagetool.UserConfig;
import org.languagetool.JLanguageTool.ParagraphHandling;
import org.languagetool.rules.Category;
import org.languagetool.rules.CategoryId;
import org.languagetool.rules.ITSIssueType;
import org.languagetool.rules.Rule;
import org.languagetool.rules.RuleMatch;
import org.languagetool.rules.RuleOption;
import org.languagetool.rules.TextLevelRule;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtOfficeTools.RemoteCheck;

/**
 * Class to run LanguageTool in LO to use a remote server
 * @since 1.0
 * @author Fred Kruse
 */
public class WtRemoteLanguageTool {
  
  private static boolean debugMode = false;   //  should be false except for testing

  private static final String BLANK = " ";
  private static final String SERVER_URL = "https://api.languagetool.org";
  private static final String PREMIUM_SERVER_URL = "https://api.languagetoolplus.com";
  private static final int MAX_LIMIT = 2147483647;
  private static final int SERVER_LIMIT = 20000;

  private final Set<String> enabledRules = new HashSet<>();
  private final Set<String> disabledRules = new HashSet<>();
  private final Set<CategoryId> disabledRuleCategories = new HashSet<>();
  private final Set<CategoryId> enabledRuleCategories = new HashSet<>();
  private final List<Rule> allRules = new ArrayList<>();
  private final List<Rule> spellingRules = new ArrayList<>();
  private final List<String> ruleValues = new ArrayList<>();
  private final Language language;
  private final Language motherTongue;
  private final RemoteLanguageTool remoteLanguageTool;
//  private final UserConfig userConfig;
//  private final boolean addSynonyms;
  private final boolean isPremium;
  private final String username;
  private final String apiKey;
//  private JLanguageTool lt = null;

  private int maxTextLength = MAX_LIMIT;
  private boolean remoteRun;
  
  public WtRemoteLanguageTool(Language language, Language motherTongue, WtConfiguration config,
                       List<Rule> extraRemoteRules, UserConfig userConfig) throws MalformedURLException {
    debugMode = WtOfficeTools.DEBUG_MODE_RM;
    this.language = language;
    this.motherTongue = motherTongue;
//    this.userConfig = userConfig;
//    addSynonyms = userConfig != null && userConfig.getLinguServices() != null && !config.noSynonymsAsSuggestions();
    String serverUrl = config.getServerUrl();
    setRuleValues(config.getConfigurableValues());
    username = config.getRemoteUsername();
    apiKey = config.getRemoteApiKey();
    isPremium = username != null && apiKey != null && config.isPremium();
    URL serverBaseUrl = new URL(serverUrl == null ? (isPremium ? PREMIUM_SERVER_URL : SERVER_URL) : serverUrl);
    remoteLanguageTool = new RemoteLanguageTool(serverBaseUrl);
    try {
      String urlParameters = "language=" + language.getShortCodeWithCountryAndVariant();
      RemoteConfigurationInfo configInfo = remoteLanguageTool.getConfigurationInfo(urlParameters);
      storeAllRules(configInfo.getRemoteRules());
      if (!isPremium) {
        maxTextLength = remoteLanguageTool.getMaxTextLength();
      }
      WtMessageHandler.printToLogFile("Server Url: " + serverBaseUrl);
      WtMessageHandler.printToLogFile("Server Limit text length: " + maxTextLength);
      remoteRun = true;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
      maxTextLength = SERVER_LIMIT;
      WtMessageHandler.printToLogFile("Server doesn't support maxTextLength, Limit text length set to: " + maxTextLength);
      remoteRun = false;
    }
  }
  
  /**
   * check a text by a remote LT server
   */
  public List<RuleMatch> check(String text, ParagraphHandling paraMode, RemoteCheck checkMode) throws Throwable {
    if (!remoteRun) {
      return null;
    }
    List<RuleMatch> ruleMatches = new ArrayList<>();
    if (text == null || text.trim().isEmpty()) {
      return ruleMatches;
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("WtRemoteLanguageTool: check: check text: " + text);
    }
    CheckConfigurationBuilder configBuilder = new CheckConfigurationBuilder(language.getShortCodeWithCountryAndVariant());
    if (isPremium) {
      configBuilder.username(username);
      configBuilder.apiKey(apiKey);
    }
    if (motherTongue != null) {
      configBuilder.setMotherTongueLangCode(motherTongue.getShortCodeWithCountryAndVariant());
    }
    if (paraMode == ParagraphHandling.ONLYPARA) {
      configBuilder.ruleValues(ruleValues);
      Set<String> tmpEnabled = new HashSet<>();
      if (checkMode == RemoteCheck.ALL || checkMode == RemoteCheck.ONLY_GRAMMAR) {
        tmpEnabled.addAll(enabledRules);
      }
      if (checkMode == RemoteCheck.ALL || checkMode == RemoteCheck.ONLY_SPELL) {
        for (Rule rule : spellingRules) {
          tmpEnabled.add(rule.getId());
        }
      }
      if (tmpEnabled.size() > 0) {
        configBuilder.enabledRuleIds(tmpEnabled.toArray(new String[0]));
        configBuilder.enabledOnly();
      }
      configBuilder.mode("textLevelOnly");
    } else {
      if (checkMode == RemoteCheck.ALL || checkMode == RemoteCheck.ONLY_GRAMMAR) {
        Set<String> tmpDisabled = new HashSet<>(disabledRules);
        if (checkMode == RemoteCheck.ALL) {
          for (Rule rule : spellingRules) {
            tmpDisabled.remove(rule.getId());
          }
        }
        configBuilder.enabledRuleIds(enabledRules.toArray(new String[0]));
        configBuilder.disabledRuleIds(tmpDisabled.toArray(new String[0]));
        configBuilder.ruleValues(ruleValues);
        configBuilder.mode("all");
      } else if (checkMode == RemoteCheck.ONLY_SPELL) {
        Set<String> tmpEnabled = new HashSet<>();
        for (Rule rule : spellingRules) {
          tmpEnabled.add(rule.getId());
        }
        if (tmpEnabled.size() > 0) {
          configBuilder.enabledRuleIds(tmpEnabled.toArray(new String[0]));
          configBuilder.enabledOnly();
        }
        configBuilder.mode("allButTextLevelOnly");
      }
    }
    configBuilder.level("default");
    CheckConfiguration remoteConfig = configBuilder.build();
    int limit;
    for (int nStart = 0; text.length() > nStart; nStart += limit) {
      String subText;
      if (text.length() <= nStart + maxTextLength) {
        subText = text.substring(nStart);
        limit = maxTextLength;
      } else {
        int nEnd = text.lastIndexOf(WtOfficeTools.END_OF_PARAGRAPH, nStart + maxTextLength) + WtOfficeTools.NUMBER_PARAGRAPH_CHARS;
        if (nEnd <= nStart) {
          nEnd = text.lastIndexOf(BLANK, nStart + maxTextLength) + 1;
          if (nEnd <= nStart) {
            nEnd = nStart + maxTextLength;
          }
        }
        subText = text.substring(nStart, nEnd);
        limit = nEnd;
        if (debugMode) {
          WtMessageHandler.printToLogFile("WtRemoteLanguageTool: check: split text (nStart/nEnd/text.length): (" 
                  + nStart + "/" + nEnd + "/" + text.length() + ")");
        }
     }
      RemoteResult remoteResult;
      try {
        remoteResult = remoteLanguageTool.check(subText, remoteConfig);
      } catch (Throwable t) {
        WtMessageHandler.printException(t);
        remoteRun = false;
        return null;
      }
      ruleMatches.addAll(toRuleMatches(text, remoteResult.getMatches(), nStart));
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("WtRemoteLanguageTool: check: number rule matches found: " + ruleMatches.size());
    }
    return ruleMatches;
  }
  
  /**
   * Get the language the check will done for
   */
  public Language getLanguage() {
    return language;
  }
  
  /**
   * Get all rules 
   */
  public List<Rule> getAllRules() {
    return allRules;
  }
  
  /**
   * true if the check should be done by a remote server
   */
  public boolean remoteRun() {
    return remoteRun;
  }
  
  /**
   * true if the rule should be ignored
   */
  private boolean ignoreRule(Rule rule) {
    Category ruleCategory = rule.getCategory();
    boolean isCategoryDisabled = (disabledRuleCategories.contains(ruleCategory.getId()) || rule.getCategory().isDefaultOff()) 
            && !enabledRuleCategories.contains(ruleCategory.getId());
    boolean isRuleDisabled = disabledRules.contains(rule.getId()) 
            || (rule.isDefaultOff() && !enabledRules.contains(rule.getId()));
    boolean isDisabled;
    if (isCategoryDisabled) {
      isDisabled = !enabledRules.contains(rule.getId());
    } else {
      isDisabled = isRuleDisabled;
    }
    return isDisabled;
  }

  /**
   * get all active rules
   */
  public List<Rule> getAllActiveRules() {
    List<Rule> rulesActive = new ArrayList<>();
    for (Rule rule : allRules) {
      if (!ignoreRule(rule)) {
        rulesActive.add(rule);
      }
    }    
    return rulesActive;
  }
  
  /**
   * get all active office rules
   */
  public List<Rule> getAllActiveOfficeRules() {
    List<Rule> rulesActive = new ArrayList<>();
    for (Rule rule : allRules) {
      if (!ignoreRule(rule) && !rule.isOfficeDefaultOff()) {
        rulesActive.add(rule);
      } else if (rule.isOfficeDefaultOn() && !disabledRules.contains(rule.getId())) {
        rulesActive.add(rule);
        enableRule(rule.getId());
      } else if (!ignoreRule(rule) && rule.isOfficeDefaultOff() && !enabledRules.contains(rule.getId())) {
        disableRule(rule.getId());
      }
    }    
    return rulesActive;
  }
  
  /**
   * Get disabled rules
   */
  public Set<String> getDisabledRules() {
    return disabledRules;
  }
  
  /**
   * Enable the rule
   */
  public void enableRule (String ruleId) {
    disabledRules.remove(ruleId);
    enabledRules.add(ruleId);
  }
  
  /**
   * Disable the rule
   */
  public void disableRule (String ruleId) {
    disabledRules.add(ruleId);
    enabledRules.remove(ruleId);
  }
  
  /**
   * Disable the category
   */
  public void disableCategory(CategoryId id) {
    disabledRuleCategories.add(id);
    enabledRuleCategories.remove(id);
  }
  
  /**
   * Set the values for rules
   */
  private void setRuleValues(Map<String, Object[]> configurableValues) {
    ruleValues.clear();
    Set<String> rules = configurableValues.keySet();
    for (String rule : rules) {
      Object[] obs = configurableValues.get(rule);
      if (obs != null && obs.length > 0) {
        String rOptions = RuleOption.objectsToString(obs);
        if (rOptions.charAt(0) == 'i') {  // for compatibility with older server versions
          rOptions = rOptions.substring(1);
        }
        String ruleValueString = rule + ":" + rOptions;
        ruleValues.add(ruleValueString);
      }
    }
  }
  
  /**
   * get synonyms for a word
   *//*
  private List<String> getSynonymsForWord(String word, LinguServices linguServices) {
    List<String> synonyms = new ArrayList<String>();
    List<String> rawSynonyms = linguServices.getSynonyms(word, language);
    for (String synonym : rawSynonyms) {
      synonym = synonym.replaceAll("\\(.*\\)", "").trim();
      if (!synonym.isEmpty() && !synonyms.contains(synonym)) {
        synonyms.add(synonym);
      }
    }
    return synonyms;
  }

  /**
   * get synonyms for a repeated word
   *//*
  private List<String> getSynonymsForToken(AnalyzedTokenReadings token, LinguServices linguServices) {
    List<String> synonyms = new ArrayList<String>();
    if(linguServices == null || token == null) {
      return synonyms;
    }
    List<AnalyzedToken> readings = token.getReadings();
    for (AnalyzedToken reading : readings) {
      String lemma = reading.getLemma();
      if (lemma != null) {
        List<String> newSynonyms = getSynonymsForWord(lemma, linguServices);
        for (String synonym : newSynonyms) {
          if (!synonyms.contains(synonym)) {
            synonyms.add(synonym);
          }
        }
      }
    }
    if(synonyms.isEmpty()) {
      synonyms = getSynonymsForWord(token.getToken(), linguServices);
    }
    return synonyms;
  }
  /**
   * get synonyms of a word
   *//*
  private List<String> getSynonyms(String word) {
    if (!addSynonyms) {
      return null;
    }
    if (lt == null) {
      lt = new JLanguageTool(language, motherTongue, null, userConfig);
    }
    List<AnalyzedSentence> analyzedSentence;
    try {
      analyzedSentence = lt.analyzeText(word);
      return getSynonymsForToken(analyzedSentence.get(0).getTokensWithoutWhitespace()[1], userConfig.getLinguServices());
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
      return null;
    }
  }
  
  /**
   * Convert a remote rule match to a LT rule match 
   */
  private RuleMatch toRuleMatch(String text, RemoteRuleMatch remoteMatch, int nOffset) throws MalformedURLException {
    Rule matchRule = null;
    for (Rule rule : allRules) {
      if (remoteMatch.getRuleId().equals(rule.getId())) {
        matchRule = rule;
      }
    }
    if (matchRule == null) {
      WtMessageHandler.printToLogFile("WARNING: Rule \"" + remoteMatch.getRuleDescription() + "(ID: " 
                                    + remoteMatch.getRuleId() + ")\" may be not supported by option panel!");
      matchRule = new RemoteRule(remoteMatch);
      allRules.add(matchRule);
    }
    RuleMatch ruleMatch = new RuleMatch(matchRule, null, remoteMatch.getErrorOffset() + nOffset, 
        remoteMatch.getErrorOffset() + remoteMatch.getErrorLength() + nOffset, remoteMatch.getMessage(), 
        remoteMatch.getShortMessage().isPresent() ? remoteMatch.getShortMessage().get() : null);
    if (remoteMatch.getUrl().isPresent()) {
      ruleMatch.setUrl(new URL(remoteMatch.getUrl().get()));
    }
    List<String> replacements = null;
    if (remoteMatch.getReplacements().isPresent()) {
      replacements = remoteMatch.getReplacements().get();
    }
    if (replacements != null && !replacements.isEmpty()) {
      ruleMatch.setSuggestedReplacements(remoteMatch.getReplacements().get());
/*      
    } else if (addSynonyms && remoteMatch.getRuleId().startsWith("STYLE_REPEATED_WORD_RULE")) {
      String word = text.substring(ruleMatch.getFromPos(), ruleMatch.getToPos());
      List<String> synonyms = getSynonyms(word);
      if (synonyms != null && !synonyms.isEmpty()) {
        ruleMatch.setSuggestedReplacements(synonyms);
      }
*/
    }
    return ruleMatch;
  }
  
  /**
   * Convert a list of remote rule matches to a list of LT rule matches
   */
  private List<RuleMatch> toRuleMatches(String text, List<RemoteRuleMatch> remoteRulematches, int nOffset) throws MalformedURLException {
    List<RuleMatch> ruleMatches = new ArrayList<>();
    if (remoteRulematches == null || remoteRulematches.isEmpty()) {
      return ruleMatches;
    }
    for (RemoteRuleMatch remoteMatch : remoteRulematches) {
      RuleMatch ruleMatch = toRuleMatch(text, remoteMatch, nOffset);
      ruleMatches.add(ruleMatch);
    }
    return ruleMatches;
  }
  
  /**
   * store all rules in a list
   */
  private void storeAllRules(List<Map<String,String>> listRuleMaps) {
    allRules.clear();
    spellingRules.clear();
    for (Map<String,String> ruleMap : listRuleMaps) {
      Rule rule;
      if (ruleMap.containsKey("isTextLevelRule")) {
        rule = new RemoteTextLevelRule(ruleMap);
      } else {
        rule = new RemoteRule(ruleMap);
      }
      if (rule.isDictionaryBasedSpellingRule()) {
        spellingRules.add(rule);
      }
      allRules.add(rule);
    }
  }
  
  public static RuleOption[] getRuleOptionsFromMap(Map<String,String> ruleMap) {
    if (ruleMap.containsKey("ruleOptions")) {
      Object o = ruleMap.get("ruleOptions");
      @SuppressWarnings("unchecked")
      List<Map<String,Object>> rOptions = (List<Map<String, Object>>) o;
      RuleOption[] ruleOptions = new RuleOption[rOptions.size()];
      for (int i = 0; i < rOptions.size(); i++) {
        String defaultType = (String) rOptions.get(i).get(RuleOption.DEFAULT_TYPE);
        String configureText = (String) rOptions.get(i).get(RuleOption.CONF_TEXT);
        Object defaultValue;
        Object minConfigurableValue;
        Object maxConfigurableValue;
        if (defaultType.equals("Integer")) {
          defaultValue = (int) rOptions.get(i).get(RuleOption.DEFAULT_VALUE);
          minConfigurableValue = (int) rOptions.get(i).get(RuleOption.MIN_CONF_VALUE);
          maxConfigurableValue = (int) rOptions.get(i).get(RuleOption.MAX_CONF_VALUE);
        } else if (defaultType.equals("Character")) {
          defaultValue = rOptions.get(i).get(RuleOption.DEFAULT_VALUE).toString().charAt(0);
          minConfigurableValue = rOptions.get(i).get(RuleOption.MIN_CONF_VALUE).toString().charAt(0);
          maxConfigurableValue = rOptions.get(i).get(RuleOption.MAX_CONF_VALUE).toString().charAt(0);
        } else if (defaultType.equals("Boolean")) {
          defaultValue = (boolean) rOptions.get(i).get(RuleOption.DEFAULT_VALUE);
          minConfigurableValue = (int) rOptions.get(i).get(RuleOption.MIN_CONF_VALUE);
          maxConfigurableValue = (int) rOptions.get(i).get(RuleOption.MAX_CONF_VALUE);
        } else if (defaultType.equals("Float")) {
          defaultValue = (float) rOptions.get(i).get(RuleOption.DEFAULT_VALUE);
          minConfigurableValue = (float) rOptions.get(i).get(RuleOption.MIN_CONF_VALUE);
          maxConfigurableValue = (float) rOptions.get(i).get(RuleOption.MAX_CONF_VALUE);
        } else if (defaultType.equals("Double")) {
          defaultValue = (double) rOptions.get(i).get(RuleOption.DEFAULT_VALUE);
          minConfigurableValue = (double) rOptions.get(i).get(RuleOption.MIN_CONF_VALUE);
          maxConfigurableValue = (double) rOptions.get(i).get(RuleOption.MAX_CONF_VALUE);
        } else {
          defaultValue = rOptions.get(i).get(RuleOption.DEFAULT_VALUE).toString();
          minConfigurableValue = rOptions.get(i).get(RuleOption.MIN_CONF_VALUE).toString();
          maxConfigurableValue = rOptions.get(i).get(RuleOption.MAX_CONF_VALUE).toString();
        }
        ruleOptions[i] = new RuleOption(defaultValue, configureText, minConfigurableValue, maxConfigurableValue);
      }
      return ruleOptions;
    } else {
      return null;
    }
  }

  /**
   * Class to define remote (sentence level) rules
   */
  static class RemoteRule extends Rule {
    
    private final String ruleId;
    private final String description;
    private final boolean hasConfigurableValue;
    private final boolean isDictionaryBasedSpellingRule;
    private final int defaultValue;
    private final int minConfigurableValue;
    private final int maxConfigurableValue;
    private final String configureText;
    private final RuleOption[] ruleOptions;
    
    RemoteRule(Map<String,String> ruleMap) {
      ruleId = ruleMap.get("ruleId");
      description = ruleMap.get("description");
      if (ruleMap.containsKey("isDefaultOff")) {
        setDefaultOff();
      }
      if (ruleMap.containsKey("isOfficeDefaultOn")) {
        setOfficeDefaultOn();
      }
      if (ruleMap.containsKey("isOfficeDefaultOff")) {
        setOfficeDefaultOff();
      }
      isDictionaryBasedSpellingRule = ruleMap.containsKey("isDictionaryBasedSpellingRule");
      ruleOptions = getRuleOptionsFromMap(ruleMap);
      if (ruleMap.containsKey("hasConfigurableValue")) {
        hasConfigurableValue = true;
        defaultValue = Integer.parseInt(ruleMap.get(RuleOption.DEFAULT_VALUE));
        minConfigurableValue = Integer.parseInt(ruleMap.get(RuleOption.MIN_CONF_VALUE));
        maxConfigurableValue = Integer.parseInt(ruleMap.get(RuleOption.MAX_CONF_VALUE));
        configureText = ruleMap.get(RuleOption.CONF_TEXT);
      } else {
        hasConfigurableValue = false;
        defaultValue = 0;
        minConfigurableValue = 0;
        maxConfigurableValue = 100;
        configureText = "";
      }
      setCategory(new Category(new CategoryId(ruleMap.get("categoryId")), ruleMap.get("categoryName")));
      setLocQualityIssueType(ITSIssueType.getIssueType(ruleMap.get("locQualityIssueType")));
    }

    RemoteRule(RemoteRuleMatch remoteMatch) {
      ruleId = remoteMatch.getRuleId();
      description = remoteMatch.getRuleDescription();
      isDictionaryBasedSpellingRule = false;
      hasConfigurableValue = false;
      defaultValue = 0;
      minConfigurableValue = 0;
      maxConfigurableValue = 100;
      configureText = "";
      ruleOptions = null;
      String categoryId = remoteMatch.getCategoryId().orElse(null);
      String categoryName = remoteMatch.getCategory().orElse(null);
      if (categoryId != null && categoryName != null) {
        setCategory(new Category(new CategoryId(categoryId), categoryName));
      }
      String locQualityIssueType = remoteMatch.getLocQualityIssueType().orElse(null);
      if (locQualityIssueType != null) {
        setLocQualityIssueType(ITSIssueType.getIssueType(locQualityIssueType));
      }
    }

    @Override
    public String getId() {
      return ruleId;
    }

    @Override
    public String getDescription() {
      return description;
    }

    @Override
    public boolean isDictionaryBasedSpellingRule() {
      return isDictionaryBasedSpellingRule;
    }

    @Override
    public RuleOption[] getRuleOptions() {
      if (ruleOptions != null) {
        return ruleOptions;
      }
      if (hasConfigurableValue) {
        RuleOption[] ruleOptions = { new RuleOption(defaultValue, configureText, minConfigurableValue, maxConfigurableValue) };
        return ruleOptions;
      }
      return null;
    }

    @Override
    public RuleMatch[] match(AnalyzedSentence sentence) {
      return null;
    }
    
  }
  
  /**
   * Class to define remote text level rules
   */
  static class RemoteTextLevelRule extends TextLevelRule {
    
    private final String ruleId;
    private final String description;
    private final boolean hasConfigurableValue;
    private final boolean isDictionaryBasedSpellingRule;
    private final int defaultValue;
    private final int minConfigurableValue;
    private final int maxConfigurableValue;
    private final int minToCheckParagraph;
    private final String configureText;
    private final RuleOption[] ruleOptions;
    
    RemoteTextLevelRule(Map<String,String> ruleMap) {
      ruleId = ruleMap.get("ruleId");
      description = ruleMap.get("description");
      if (ruleMap.containsKey("isDefaultOff")) {
        setDefaultOff();
      }
      if (ruleMap.containsKey("isOfficeDefaultOn")) {
        setOfficeDefaultOn();
      }
      if (ruleMap.containsKey("isOfficeDefaultOff")) {
        setOfficeDefaultOff();
      }
      isDictionaryBasedSpellingRule = ruleMap.containsKey("isDictionaryBasedSpellingRule");
      ruleOptions = getRuleOptionsFromMap(ruleMap);
      if (ruleMap.containsKey("hasConfigurableValue")) {
        hasConfigurableValue = true;
        defaultValue = Integer.parseInt(ruleMap.get(RuleOption.DEFAULT_VALUE));
        minConfigurableValue = Integer.parseInt(ruleMap.get(RuleOption.MIN_CONF_VALUE));
        maxConfigurableValue = Integer.parseInt(ruleMap.get(RuleOption.MAX_CONF_VALUE));
        configureText = ruleMap.get(RuleOption.CONF_TEXT);
      } else {
        hasConfigurableValue = false;
        defaultValue = 0;
        minConfigurableValue = 0;
        maxConfigurableValue = 100;
        configureText = "";
      }
      minToCheckParagraph = Integer.parseInt(ruleMap.get("minToCheckParagraph"));
      setCategory(new Category(new CategoryId(ruleMap.get("categoryId")), ruleMap.get("categoryName")));
      setLocQualityIssueType(ITSIssueType.getIssueType(ruleMap.get("locQualityIssueType")));
    }

    @Override
    public String getId() {
      return ruleId;
    }

    @Override
    public String getDescription() {
      return description;
    }

    @Override
    public boolean isDictionaryBasedSpellingRule() {
      return isDictionaryBasedSpellingRule;
    }

    @Override
    public RuleOption[] getRuleOptions() {
      if (ruleOptions != null) {
        return ruleOptions;
      }
      if (hasConfigurableValue) {
        RuleOption[] ruleOptions = { new RuleOption(defaultValue, configureText, minConfigurableValue, maxConfigurableValue) };
        return ruleOptions;
      }
      return null;
    }

    @Override
    public RuleMatch[] match(List<AnalyzedSentence> sentences) {
      return null;
    }

    @Override
    public int minToCheckParagraph() {
      return minToCheckParagraph;
    }
    
  }
  
}
