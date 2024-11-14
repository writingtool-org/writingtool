/* WritingTool, a LibreOffice Extension based on LanguageTool
 * Copyright (C) 2024 Fred Kruse (https://fk-es.de)
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
package org.writingtool.stylestatistic;

import java.awt.Color;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import org.languagetool.rules.AbstractStatisticSentenceStyleRule;
import org.languagetool.rules.AbstractStatisticStyleRule;
import org.languagetool.rules.AbstractStyleTooOftenUsedWordRule;
import org.languagetool.rules.TextLevelRule;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtVersionInfo;

/**
 * Statistical Analyzes Configuration 
 * @since 6.2
 * @author Fred Kruse
 */
public class WtStatAnConfiguration {
  
  private final boolean WITHOUT_DIRECT_SPEECH_DEFAULT = false;
  private final boolean SHOW_ALLPARAGRAPHS_DEFAULT = false;
  private final short UNDERLINE_TYPE_DEFAULT = 0;
  private final Color UNDERLINE_COLOR_DEFAULT = new Color(75, 0, 255);
  
  private final String UNDERLINE_TYPE_PROP = "UnderlineType";
  private final String UNDERLINE_COLOR_PROP = "UnderlineColor";
  private final String SHOW_ALL_PARAGRAPH_PROP = "ShowAllParagraphs";
  
  private final List<TextLevelRule> rules;
  private final Map<String, Boolean> withoutDirectSpeech = new HashMap<String, Boolean>();
  private final Map<String, Integer> levelStep = new HashMap<String, Integer>();
  private final Map<String, List<String>> excludedWords = new HashMap<String, List<String>>();
  private short underlineType = UNDERLINE_TYPE_DEFAULT;
  private Color underlineColor = UNDERLINE_COLOR_DEFAULT;
  private boolean showAllParagraphs = SHOW_ALLPARAGRAPHS_DEFAULT;
  private boolean showAdditionalOptions = false;
  
  public WtStatAnConfiguration(List<TextLevelRule> rules) throws Throwable {
    this.rules = rules;
    loadConfiguration();
  }
  
  public boolean isWithoutDirectSpeech(TextLevelRule rule) {
    if (!withoutDirectSpeech.containsKey(rule.getId())) {
      return false;
    }
    return withoutDirectSpeech.get(rule.getId());
  }
  
  public void setWithoutDirectSpeech(TextLevelRule rule, boolean wDS) {
    withoutDirectSpeech.put(rule.getId(), wDS);
  }
  
  public int getLevelStep(TextLevelRule rule) {
    if (!levelStep.containsKey(rule.getId())) {
      return 0;
    }
    return levelStep.get(rule.getId());
  }
  
  public void setLevelStep(TextLevelRule rule, int step) {
    levelStep.put(rule.getId(), step);
  }

  public List<String> getExcludedWords(TextLevelRule rule) {
    List<String> exWords = excludedWords.get(rule.getId());
    if (exWords == null) {
      return new ArrayList<>();
    }
    return exWords;
  }

  public void setAllExcludedWords(TextLevelRule rule, List<String> words) {
    excludedWords.put(rule.getId(), words);
  }

  public void removeAllExcludedWords(TextLevelRule rule) {
    List<String> exWords = excludedWords.get(rule.getId());
    if (exWords != null) {
      exWords.clear();
      excludedWords.put(rule.getId(), exWords);
    }
  }

  public void addExcludedWord(TextLevelRule rule, String word) {
    List<String> words = excludedWords.get(rule.getId());
    if (words == null) {
      words = new ArrayList<>();
    }
    words.add(word);
    excludedWords.put(rule.getId(), words);
  }
  
  public short getUnderlineType() {
    return underlineType;
  }
  
  public void setUnderlineType(short underlineType) {
    this.underlineType = underlineType;
  }
  
  public boolean showAllParagraphs() {
    return showAllParagraphs;
  }
  
  public void setShowAllParagraphs(boolean showAllParagraphs) {
    this.showAllParagraphs = showAllParagraphs;
  }
  
  public boolean showAdditionalOptions() {
    return this.showAdditionalOptions;
  }
  
  public void setShowAdditionalOptions(boolean showAdditionalOptions) {
    this.showAdditionalOptions = showAdditionalOptions;
  }
  
  public Color getUnderlineColor() {
    return underlineColor;
  }
  
  public void setDefaultUnderlineColor() {
    underlineColor = UNDERLINE_COLOR_DEFAULT;
  }
  
  public void setUnderlineColor(Color underlineColor) {
    this.underlineColor = underlineColor;
  }
  
  private String createLevelRuleProperties(String ruleId) throws Throwable {
    return Boolean.toString(withoutDirectSpeech.get(ruleId)) + ";" + Integer.toString(levelStep.get(ruleId));
  }
  
  private String createUsedWordRuleProperties(String ruleId) throws Throwable {
    String txt = Boolean.toString(withoutDirectSpeech.get(ruleId)) + ";" + Integer.toString(levelStep.get(ruleId)) + ";";
    List<String> excWords = excludedWords.get(ruleId);
    for (int i = 0; i < excWords.size(); i++) {
      txt += excWords.get(i);
      if (i < excWords.size() - 1) {
        txt += ",";
      }
    }
    return txt;
  }
  
  private int getDefaultStep(TextLevelRule rule) {
    int defaultStep;
    if (rule instanceof AbstractStatisticSentenceStyleRule) {
      defaultStep = (int) ((((int)((AbstractStatisticSentenceStyleRule) rule).getRuleOptions()[0].getDefaultValue() - 1) / 3.) + 0.5);
    } else if (rule instanceof AbstractStatisticStyleRule) {
      defaultStep = (int) ((((int)((AbstractStatisticStyleRule) rule).getRuleOptions()[0].getDefaultValue() - 1) / 3.) + 0.5);
    } else if (rule instanceof AbstractStyleTooOftenUsedWordRule) {
      defaultStep = (int) ((((int)((AbstractStyleTooOftenUsedWordRule) rule).getRuleOptions()[0].getDefaultValue() - 1) / 3.) + 0.5);
    } else {
      return -1;
    }
    return defaultStep;
  }
  
  private void setLevelRuleProperties(TextLevelRule rule, String propString) throws Throwable {
    int defaultStep = getDefaultStep(rule);
    if (defaultStep < 0) {
      return;
    }
    if (propString == null) {
      withoutDirectSpeech.put(rule.getId(), false);
      levelStep.put(rule.getId(), defaultStep);
    } else {
      String[] props = propString.split(";");
      if (props.length > 0 && !props[0].isEmpty()) {
        withoutDirectSpeech.put(rule.getId(), Boolean.parseBoolean(props[0]));
      } else {
        withoutDirectSpeech.put(rule.getId(), false);
      }
      if (props.length > 1 && props[1] != null && !props[1].isEmpty()) {
        levelStep.put(rule.getId(), Integer.parseInt(props[1]));
      } else {
        levelStep.put(rule.getId(), defaultStep);
      }
    }
  }
  
  private void setUsedWordRuleProperties(TextLevelRule rule, String propString) throws Throwable {
    try {
      AbstractStyleTooOftenUsedWordRule usedWordRule = (AbstractStyleTooOftenUsedWordRule) rule;
      List<String> excWords = new ArrayList<String>();
      int defaultStep = getDefaultStep(rule);
      if (propString == null) {
  //      MessageHandler.printToLogFile("Rule: " + usedWordRule.getId() + "; PropString == null");
  //      MessageHandler.printToLogFile("DefaultDirectSpeach(): " + usedWordRule.getDefaultDirectSpeach() + "; excWords.size: " + excWords.size());
        withoutDirectSpeech.put(usedWordRule.getId(), false);
        levelStep.put(rule.getId(), defaultStep);
        excludedWords.put(usedWordRule.getId(), excWords);
      } else {
        String[] props = propString.split(";");
        if (props.length > 0 && !props[0].isEmpty()) {
          withoutDirectSpeech.put(usedWordRule.getId(), Boolean.parseBoolean(props[0]));
        } else {
          withoutDirectSpeech.put(usedWordRule.getId(), false);
        }
        if (props.length > 1 && props[1] != null && !props[1].isEmpty()) {
          levelStep.put(rule.getId(), Integer.parseInt(props[1]));
        } else {
          levelStep.put(rule.getId(), defaultStep);
        }
        if (props.length > 2 && props[2] != null && !props[2].isEmpty()) {
          String[] words = props[2].split(",");
          for (String word : words) {
            excWords.add(word);
          }
        }
        excludedWords.put(usedWordRule.getId(), excWords);
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }
  
  public void saveConfiguration() throws Throwable {
    Properties props = new Properties();
    for (TextLevelRule rule : rules) {
      if (rule instanceof AbstractStyleTooOftenUsedWordRule) {
        if (withoutDirectSpeech.get(rule.getId()) != WITHOUT_DIRECT_SPEECH_DEFAULT 
            || levelStep.get(rule.getId()) != getDefaultStep(rule) || excludedWords.get(rule.getId()) != null) {
          props.setProperty(rule.getId(), createUsedWordRuleProperties(rule.getId()));
        }
      } else if (WtLevelRule.hasStatisticalOptions(rule)) {
        if (withoutDirectSpeech.get(rule.getId()) != WITHOUT_DIRECT_SPEECH_DEFAULT 
            || levelStep.get(rule.getId()) != getDefaultStep(rule)) {
          props.setProperty(rule.getId(), createLevelRuleProperties(rule.getId()));
        }
      }
    }
    if (underlineType != UNDERLINE_TYPE_DEFAULT) {
      props.setProperty(UNDERLINE_TYPE_PROP, Short.toString(underlineType));
    }
    if (!underlineColor.equals(UNDERLINE_COLOR_DEFAULT)) {
      props.setProperty(UNDERLINE_COLOR_PROP, Integer.toString(underlineColor.getRGB()));
    }
    if (showAllParagraphs != SHOW_ALLPARAGRAPHS_DEFAULT) {
      props.setProperty(SHOW_ALL_PARAGRAPH_PROP, Boolean.toString(showAllParagraphs));
      
    }
    try (FileOutputStream fos = new FileOutputStream(WtOfficeTools.getStatisticalConfigFilePath())) {
      props.store(fos, WtOfficeTools.WT_NAME + " statistical analyzes configuration (" + WtVersionInfo.getWtInformation() + ")");
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    }
  }
  
  void loadConfiguration() throws Throwable {
    Properties props = new Properties();
    try (FileInputStream fis = new FileInputStream(WtOfficeTools.getStatisticalConfigFilePath())) {
      props.load(fis);
    } catch (Throwable e) {
      WtMessageHandler.printException(e);
    }
    withoutDirectSpeech.clear();
    levelStep.clear();
    excludedWords.clear();
    for (TextLevelRule rule : rules) {
      if (rule instanceof AbstractStyleTooOftenUsedWordRule) {
        setUsedWordRuleProperties(rule, props.getProperty(rule.getId()));
      } else {
        boolean hasOptions = WtLevelRule.hasStatisticalOptions(rule);
//        MessageHandler.printToLogFile("Rule: " + rule.getId() + "; has options: " + hasOptions);
        if (hasOptions) {
          setLevelRuleProperties(rule, props.getProperty(rule.getId()));
        }
      }
    }
    String propString = props.getProperty(UNDERLINE_TYPE_PROP);
    if (propString != null) {
      underlineType = Short.parseShort(propString);
    }
    propString = props.getProperty(UNDERLINE_COLOR_PROP);
    if (propString != null) {
      underlineColor = Color.decode(propString);
    }
    propString = props.getProperty(SHOW_ALL_PARAGRAPH_PROP);
    if (propString != null) {
      showAllParagraphs = Boolean.parseBoolean(propString);
    }
    WtMessageHandler.printToLogFile("Config load done");
  }

}
