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
package org.writingtool.aisupport;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.ResourceBundle;
import java.util.Set;

import org.languagetool.AnalyzedSentence;
import org.writingtool.WtLinguisticServices;
import com.sun.star.lang.Locale;

/**
 * AI detection rule 
 * analyzes output of AI  
 * @since 1.0
 * @author Fred Kruse
 */
public class WtAiDetectionRule_de extends WtAiDetectionRule {
  
  private static final String CONFUSION_FILE_1 = "confusion_set_candidates.txt";
  private static final String CONFUSION_FILE_2 = "confusion_sets.txt";
  private static final String WT_CONFUSION_FILE = "confusion_sets.txt";

  private static Map<String, Set<String>> confusionWords = null;
  private static Map<String, String> noneConfusionWords = null;

  WtAiDetectionRule_de(String aiResultText, List<AnalyzedSentence> analyzedAiResult, String paraText,
      WtLinguisticServices linguServices, Locale locale, ResourceBundle messages, int showStylisticHints) throws Throwable {
    super(aiResultText, analyzedAiResult, paraText, linguServices, locale, messages, showStylisticHints);
    if (confusionWords == null) {
      confusionWords = WtAiConfusionPairs.getConfusionWordMap(locale, CONFUSION_FILE_1);
      confusionWords = WtAiConfusionPairs.getConfusionWordMap(locale, CONFUSION_FILE_2, confusionWords);
      confusionWords = WtAiConfusionPairs.getWtConfusionWordMap(locale, WT_CONFUSION_FILE, confusionWords);
    }
    if (noneConfusionWords == null) {
      noneConfusionWords = getNoneConfusionWords();
    }
  }
  
  private Map<String, String> getNoneConfusionWords() {
    Map<String, String> noneConfusionWords = new HashMap<>();
    noneConfusionWords.put("die", "sie");
    noneConfusionWords.put("Die", "Sie");
    return noneConfusionWords;
  }
  
  @Override
  public String getLanguage() {
    return "de";
  }

  /**
   * Set Exceptions to set Color for specific Languages
   */
  @Override
  public boolean isNoneHintException(int nParaStart, int nParaEnd, 
      int nResultStart, int nResultEnd, List<WtAiToken> paraTokens, List<WtAiToken> resultTokens) throws Throwable {
//    WtMessageHandler.printToLogFile("isHintException in: de" 
//        + ", paraToken: " + paraToken.getToken() + ", resultToken: " + resultToken.getToken());
    if (nParaStart == nParaEnd && nResultStart == nResultEnd) {
      String pToken = paraTokens.get(nParaStart).getToken();
      String rToken = resultTokens.get(nResultStart).getToken();
      for (String s : noneConfusionWords.keySet()) {
        if (pToken.equals(s)) {
          if (rToken.equals(noneConfusionWords.get(s))) {
            return true;
          }
        }
      }
    }
    return false;   
  }

  /**
   * Set Exceptions to set Color for specific Languages
   */
  @Override
  public boolean isHintException(int nParaStart, int nParaEnd, 
      int nResultStart, int nResultEnd, List<WtAiToken> paraTokens, List<WtAiToken> resultTokens) throws Throwable {
//    WtMessageHandler.printToLogFile("isHintException in: de" 
//        + ", paraToken: " + paraToken.getToken() + ", resultToken: " + resultToken.getToken());
    if (nParaStart == nParaEnd) {
      String pToken = paraTokens.get(nParaStart).getToken();
      if (nResultStart == nResultEnd) {
        String rToken = resultTokens.get(nResultStart).getToken();
        if ("dass".equals(pToken)  || "dass".equals(rToken)) {
          return true;
        } else if (pToken.equalsIgnoreCase(rToken)) {
          return true;
        } else if (isConfusionPair(pToken, rToken)) {
          return true;
        }
      } else if (nResultStart == nResultEnd - 1) {
        if ((",".equals(resultTokens.get(nResultStart).getToken()) 
                && (pToken.equals(resultTokens.get(nResultEnd).getToken()) || "dass".equals(resultTokens.get(nResultEnd).getToken())))
            || pToken.equals(resultTokens.get(nResultStart).getToken() + resultTokens.get(nResultEnd).getToken())) {
          return true;
        }
      }
    } else if (nResultStart == nResultEnd) {
      if (nParaStart == nParaEnd - 1) {
        String rToken = resultTokens.get(nResultStart).getToken();
        if (rToken.equals(paraTokens.get(nParaStart).getToken() + paraTokens.get(nParaEnd).getToken())) {
          return true;
        }
      }
    }
    return false;   
  }

  /**
   * If tokens (from start to end) contains sToken return true 
   * else false
   */
  private boolean containToken(String sToken, int start, int end, List<WtAiToken> tokens) throws Throwable {
    for (int i = start; i <= end; i++) {
      if (sToken.equals(tokens.get(i).getToken())) {
        return true;
      }
    }
    return false;
  }

  /**
   * Set language specific exceptions to handle change as a match
   * @throws Throwable 
   */
  @Override
  public boolean isMatchException(int nParaStart, int nParaEnd,
      int nResultStart, int nResultEnd, List<WtAiToken> paraTokens, List<WtAiToken> resultTokens) throws Throwable {
    if (nResultStart < 0 || nResultStart >= resultTokens.size() - 1) {
      return false;
    }
    if ((resultTokens.get(nResultStart).getToken().equals(",") && nResultStart < resultTokens.size() - 1 
          && (resultTokens.get(nResultStart + 1).getToken().equals("und") || resultTokens.get(nResultStart + 1).getToken().equals("oder")))
        || (resultTokens.get(nResultStart + 1).getToken().equals(",") && nResultStart < resultTokens.size() - 2 
            && (resultTokens.get(nResultStart + 2).getToken().equals("und") 
                || resultTokens.get(nResultStart + 2).getToken().equals("oder")))) {
      return true;
    }
    if (nParaStart < paraTokens.size() - 1
        && isQuote(paraTokens.get(nParaStart).getToken()) 
        && ",".equals(paraTokens.get(nParaStart + 1).getToken())
        && !containToken(paraTokens.get(nParaStart).getToken(), nResultStart, nResultEnd, resultTokens)) {
      return true;
    }
    if(nParaStart == nParaEnd && nResultStart == nResultEnd 
        && !paraTokens.get(nParaStart).isNonWord() && resultTokens.get(nResultStart).getToken().contains("-")
        && resultTokens.get(nResultStart).getToken().replace("-", "").equalsIgnoreCase(paraTokens.get(nParaStart).getToken())) {
      return true;
    }
    return false;   
  }
  
  private boolean isConfusionPair(String wPara, String wResult) {
    if (confusionWords.containsKey(wPara) && confusionWords.get(wPara).contains(wResult)) {
      return true;
    }
    if (confusionWords.containsKey(wResult) && confusionWords.get(wResult).contains(wPara)) {
      return true;
    }
    return false;
  }



}
