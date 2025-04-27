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

import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.ResourceBundle;
import java.util.regex.Pattern;

import org.languagetool.AnalyzedSentence;
import org.languagetool.AnalyzedToken;
import org.languagetool.AnalyzedTokenReadings;
import org.languagetool.Tag;
import org.languagetool.rules.Categories;
import org.languagetool.rules.ITSIssueType;
import org.languagetool.rules.RuleMatch;
import org.languagetool.rules.RuleMatch.Type;
import org.writingtool.WtLinguisticServices;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.languagetool.rules.TextLevelRule;

import com.sun.star.lang.Locale;

/**
 * Rule to detect errors by a AI API
 * @since 1.0
 * @author Fred Kruse
 */
public class WtAiDetectionRule extends TextLevelRule {

  public static final int STYLE_HINT_LIMIT = 25;  //  Limit of changed tokens in a sentence in percent (after that a style hint is assumed)
  
  private static final Pattern QUOTES = Pattern.compile("[\"“”“„»«]");
  private static final Pattern SINGLE_QUOTES = Pattern.compile("[‚‘’'›‹]");
  private static final Pattern PUNCTUATION = Pattern.compile("[,\\.!\\?:]");
  private static final Pattern OPENING_BRACKETS = Pattern.compile("[\\{\\(\\[]");

  private int debugMode = WtOfficeTools.DEBUG_MODE_AI;   //  should be false except for testing

  protected final String ruleMessage;
  protected final String ruleMessageOtherPunctuation;
  protected final String ruleMessageMissingPunctuation;
  protected final String ruleMessageUnnecessaryPunctuation;
  protected final String ruleMessageGrammar;
  protected final String ruleMessageWordConfusion;
  protected final String ruleMessageMisspelling;
  
  private final ResourceBundle messages;
  private final String aiResultText;
  private final String paraText;
  private final List<AnalyzedSentence> analyzedAiResult;
  private final int showStylisticHints;
  private final WtLinguisticServices linguServices;
  private final Locale locale;

  
  WtAiDetectionRule(String aiResultText, List<AnalyzedSentence> analyzedAiResult, String paraText,
      WtLinguisticServices linguServices, Locale locale, ResourceBundle messages, int showStylisticHints) {
    this.aiResultText = aiResultText;
    this.paraText = paraText;
    this.analyzedAiResult = analyzedAiResult;
    this.messages = messages;
    this.showStylisticHints = showStylisticHints;
    this.linguServices = linguServices;
    this.locale = locale;
    ruleMessage = messages.getString("loAiRuleMessage");
    ruleMessageOtherPunctuation = messages.getString("loAiRuleMessageOtherPunctuation");
    ruleMessageMissingPunctuation = messages.getString("loAiRuleMessageMissingPunctuation");
    ruleMessageUnnecessaryPunctuation = messages.getString("loAiRuleMessageUnnecessaryPunctuation");
    ruleMessageGrammar = messages.getString("loAiRuleMessageGrammar");
    ruleMessageWordConfusion = messages.getString("loAiRuleMessageWordConfusion");
    ruleMessageMisspelling = messages.getString("loAiRuleMessageMisspelling");
    
    setCategory(Categories.STYLE.getCategory(messages));
    setLocQualityIssueType(ITSIssueType.Grammar);
    setTags(Collections.singletonList(Tag.picky));
    
    if (debugMode > 1) {
      WtMessageHandler.printToLogFile("AiDetectionRule: showStylisticHints: " + showStylisticHints);
    }
  }

  /**
   * Override this function for specific Languages
   * zzz -> no specific language
   */
  public String getLanguage() {
    return "zzz";
  }
  
  private boolean isIgnoredToken(String paraToken, String resultToken) throws Throwable {
    if (QUOTES.matcher(resultToken).matches() || SINGLE_QUOTES.matcher(resultToken).matches()) {
      return QUOTES.matcher(paraToken).matches() || SINGLE_QUOTES.matcher(paraToken).matches();
    }
    if (resultToken.equals("-")) {
      return paraToken.equals("–");
    }
    return false;
  }

  public boolean isQuote(String token) throws Throwable {
    return (QUOTES.matcher(token).matches() || SINGLE_QUOTES.matcher(token).matches());
  }

  @Override
  public RuleMatch[] match(List<AnalyzedSentence> sentences) throws IOException {
    try {
      List<RuleMatch> matches = new ArrayList<>();
      List<AiRuleMatch> tmpMatches = new ArrayList<>();
      List<WtAiToken> paraTokens = new ArrayList<>();
      List<Integer> sentenceEnds = new ArrayList<>();
      int nRuleTokens = 0;
      int nSentence = 0;
  //    int lastSentenceStart = 0;
      int lastResultStart = 0;
      int pos = 0;
      int sEnd = 0;
      int sugStart = 0;
      int sugEnd = 0;
      int nResultTokenStart = 0;
      int nResultTokenEnd = 0;
      boolean mergeSentences = false;
      for (AnalyzedSentence sentence : sentences) {
        AnalyzedTokenReadings[] tokens = sentence.getTokensWithoutWhitespace();
        for (int i = 1; i < tokens.length; i++) {
          paraTokens.add(new WtAiToken(tokens[i], pos, sentence));
          sEnd++;
        }
        pos += sentence.getCorrectedTextLength();
        sentenceEnds.add(sEnd);
      }
      List<WtAiToken> resultTokens = new ArrayList<>();
      pos = 0;
      for (AnalyzedSentence sentence : analyzedAiResult) {
        AnalyzedTokenReadings[] tokens = sentence.getTokensWithoutWhitespace();
        for (int i = 1; i < tokens.length; i++) {
          resultTokens.add(new WtAiToken(tokens[i], pos, sentence));
        }
        pos += sentence.getCorrectedTextLength();
      }
      int i;
      int j = 0;
      for (i = 0; i < paraTokens.size() && j < resultTokens.size(); i++) {
        if (!paraTokens.get(i).getToken().equals(resultTokens.get(j).getToken()) 
            && !isIgnoredToken(paraTokens.get(i).getToken(), resultTokens.get(j).getToken())) {
          if ((i == 0 && ("{".equals(resultTokens.get(j).getToken()) || QUOTES.matcher(resultTokens.get(j).getToken()).matches()))
              && j + 1 < resultTokens.size() && paraTokens.get(i).getToken().equals(resultTokens.get(j + 1).getToken())) {
            j += 2;
            continue;
          }
          if (isQuote(paraTokens.get(i).getToken())
              && i + 1 < paraTokens.size() && paraTokens.get(i + 1).getToken().equals(resultTokens.get(j).getToken())) {
            continue;
          }
          int posStart = paraTokens.get(i).getStartPos();
          int nParaTokenStart = i;
          int nParaTokenEnd = 0;
          nResultTokenStart = 0;
          nResultTokenEnd = 0;
          AnalyzedSentence sentence = paraTokens.get(i).getSentence();
          int posEnd = 0;
          String suggestion = null;
          WtAiToken singleWordToken = null;
          boolean endFound = false;
          for (int n = 1; !endFound && i + n < paraTokens.size() && j + n < resultTokens.size(); n++) {
            for (int i1 = i + n; !endFound && i1 >= i; i1--) {
              for(int j1 = j + n; j1 >= j; j1--) {
                if (paraTokens.get(i1).getToken().equals(resultTokens.get(j1).getToken()) 
                    || isIgnoredToken(paraTokens.get(i1).getToken(), resultTokens.get(j1).getToken())) {
                  endFound = true;
                  if (i1 - 1 < i) {
                    // suggest to insert some words
                    if (i > 0) {
                      //  add them to the word before
                      posStart = paraTokens.get(i - 1).getStartPos();
                      posEnd = paraTokens.get(i1 - 1).getEndPos();
                      sugStart = resultTokens.get(j - 1).getStartPos();
                      sugEnd = resultTokens.get(j1 - 1).getEndPos();
                      singleWordToken = j == j1 ? resultTokens.get(j1 - 1) : null;
                      nParaTokenStart = i - 1;
                      nParaTokenEnd = i1 - 1;
                      nResultTokenStart = j - 1;
                      nResultTokenEnd = j1 -1;
                    } else {
                      //  add them to the word after (i == 0)
                      posEnd = paraTokens.get(i1).getEndPos();
                      nParaTokenEnd = i1;
                      if (j < 1) {
                        j = 1;
                      }
                      if (j1 < 0) {
                        j1 = 0;
                      }
                      sugStart = resultTokens.get(j - 1).getStartPos();
                      sugEnd = resultTokens.get(j1).getEndPos();
                      nResultTokenStart = j - 1;
                      nResultTokenEnd = j1;
                      singleWordToken = j - 1 == j1 ? resultTokens.get(j1) : null;
                    }
                  } else {
                    //  suggest to replace or delete some tokens
                    posEnd = paraTokens.get(i1 - 1).getEndPos();
                    nParaTokenEnd = i1 - 1;
                    if (j < 0) {
                      j = 0;
                    }
                    if (j1 < 1) {
                      j1 = 1;
                    }
                    if (j <= j1 - 1) {
                      //  replace one or more tokens
                      nResultTokenStart = j;
                      nResultTokenEnd = j1 -1;
                      if (j > 0 && PUNCTUATION.matcher(paraTokens.get(i).getToken()).matches() &&
                          !PUNCTUATION.matcher(resultTokens.get(j).getToken()).matches()) {
                        //  replace a punctuation by a word
                        sugStart = resultTokens.get(j - 1).getEndPos();
                      } else {
                        sugStart = resultTokens.get(j).getStartPos();
                      }
                      sugEnd = resultTokens.get(j1 - 1).getEndPos();
                      singleWordToken = j == j1 - 1 ? resultTokens.get(j) : null;
                      if (i > 0 && PUNCTUATION.matcher(resultTokens.get(j).getToken()).matches() &&
                          !PUNCTUATION.matcher(paraTokens.get(i).getToken()).matches()) {
                        //  replace a word by a punctuation
                        posStart = paraTokens.get(i - 1).getEndPos();
                      }
                    } else {
                      //  delete some tokens
                      if (i > 0 && !PUNCTUATION.matcher(paraTokens.get(i - 1).getToken()).matches()) {
                        posStart = paraTokens.get(i - 1).getEndPos();
                      } else {
                        posEnd = paraTokens.get(i1).getStartPos();
                      }
                      sugStart = resultTokens.get(j1).getEndPos();
                      sugEnd = resultTokens.get(j1).getEndPos();
                      nResultTokenStart = j1;
                      nResultTokenEnd = j1 -1;
                    }
                  }
                  nRuleTokens += (i1 - i + 1);
                  j = j1;
                  i = i1;
                  break;
                }
              }
            }
          }
          if (!endFound) {
            //  replace all tokens from i to end of paragraph by the suggestion from j to end of suggestion
            posEnd = paraTokens.get(paraTokens.size() - 1).getEndPos();
            nParaTokenEnd = paraTokens.size() - 1;
            sugStart = resultTokens.get(j).getStartPos();
            sugEnd = resultTokens.get(resultTokens.size() - 1).getEndPos();
            nResultTokenStart = j;
            nResultTokenEnd = resultTokens.size() - 1;
            nRuleTokens += (paraTokens.size() - i);
            j = resultTokens.size() - 1;
          }
          suggestion = sugStart >= sugEnd ? "" : aiResultText.substring(sugStart, sugEnd);
          if (!isEndQoute(j, resultTokens) && isCorrectSuggestion(suggestion, singleWordToken)
              && !isMatchException(nParaTokenStart, nParaTokenEnd, nResultTokenStart, nResultTokenEnd, paraTokens, resultTokens)) {
            RuleMatch ruleMatch = new RuleMatch(this, sentence, posStart, posEnd, ruleMessage);
            ruleMatch.addSuggestedReplacement(suggestion);
            ruleMatch.setType(Type.Hint);
            setType(nParaTokenStart, nParaTokenEnd, nResultTokenStart, nResultTokenEnd, paraTokens, resultTokens, ruleMatch);
            tmpMatches.add(new AiRuleMatch(ruleMatch, sugStart, sugEnd, nParaTokenStart, nParaTokenEnd, nResultTokenStart, nResultTokenEnd));
            if (debugMode > 1) {
              WtMessageHandler.printToLogFile("AiDetectionRule: match: found: start: " + posStart + ", end: " + posEnd
                  + ", suggestion: " + suggestion);
              WtMessageHandler.printToLogFile("AiDetectionRule: match: Message: " + ruleMatch.getMessage());
            }
          } else if (debugMode > 1) {
            WtMessageHandler.printToLogFile("AiDetectionRule: match: not correct spell: locale: " + WtOfficeTools.localeToString(locale)
                + ", suggestion: " + suggestion);
          }
        }
        j++;
        if (i >= sentenceEnds.get(nSentence)) {
//          WtMessageHandler.printToLogFile("AiDetectionRule: match: resultTokens.get(j - 1).getToken(): " 
//                + resultTokens.get(nResultTokenStart).getToken());          
          if (i > sentenceEnds.get(nSentence) || !PUNCTUATION.matcher(resultTokens.get(nResultTokenEnd).getToken()).matches()) {
            //  merge sentence only if there is no punctuation change
            mergeSentences = true;
          }
          while (i >= sentenceEnds.get(nSentence)) {
            nSentence++;
          }
        }
        if (i == sentenceEnds.get(nSentence) - 1) {
          if (nRuleTokens > 0) {
            int nSenTokens = nSentence == 0 ? sentenceEnds.get(nSentence) : sentenceEnds.get(nSentence) - sentenceEnds.get(nSentence - 1);
            if (mergeSentences || styleHintAssumed(nRuleTokens, nSenTokens, tmpMatches, paraTokens, resultTokens)) {
              if (showStylisticHints == 2 && !tmpMatches.isEmpty()) {
                int startPos = tmpMatches.get(0).ruleMatch.getFromPos();
                int endPos = tmpMatches.get(tmpMatches.size() - 1).ruleMatch.getToPos();
                RuleMatch ruleMatch = new RuleMatch(this, null, startPos, endPos, ruleMessage);
                int suggestionStart = tmpMatches.get(0).suggestionStart;
                int suggestionEnd = tmpMatches.get(tmpMatches.size() - 1).suggestionEnd;
                String suggestion = aiResultText.substring(suggestionStart, suggestionEnd);
                ruleMatch.addSuggestedReplacement(suggestion);
                ruleMatch.setType(Type.Other);
                matches.add(ruleMatch);
                if (debugMode > 1) {
                  WtMessageHandler.printToLogFile("AiDetectionRule: match(1): Stylistic hint: suggestion: " + suggestion);
                  WtMessageHandler.printToLogFile("AiDetectionRule: match: mergeSentences: " + mergeSentences 
                      + ", styleHintAssumed: " + styleHintAssumed(nRuleTokens, nSenTokens, tmpMatches, paraTokens, resultTokens));
                }
              }
              mergeSentences = false;
            } else {
              addAllRuleMatches(matches, tmpMatches, resultTokens);
              if (debugMode > 1) {
                WtMessageHandler.printToLogFile("AiDetectionRule: match: add matches(1): " + tmpMatches.size()
                + ", total: " + matches.size());
              }
            }
          }
          tmpMatches.clear();
          nRuleTokens = 0;
          nSentence++;
  //        lastSentenceStart = i + 1;
          lastResultStart = j;
        }
      }
//      WtMessageHandler.printToLogFile("i = " + i + ", j = " + j + ", resultTokens.size() = " + resultTokens.size() 
//        + ", resultTokens(resultTokens.size() - 2): " + resultTokens.get(resultTokens.size() - 2).getToken()
//        + ", resultTokens(resultTokens.size() - 1): " + resultTokens.get(resultTokens.size() - 1).getToken()
//        + ", paraTokens.get(paraTokens.size() - 1): " + paraTokens.get(paraTokens.size() - 1).getToken()); 
      if (PUNCTUATION.matcher(resultTokens.get(resultTokens.size() - 1).getToken()).matches()
          && !PUNCTUATION.matcher(paraTokens.get(paraTokens.size() - 1).getToken()).matches()
          && paraTokens.get(paraTokens.size() - 1).getToken().equals(resultTokens.get(resultTokens.size() - 2).getToken())) {
        //  missing punctuation at paragraph end
        RuleMatch ruleMatch = new RuleMatch(this, null, paraTokens.get(paraTokens.size() - 1).getStartPos(), 
            paraTokens.get(paraTokens.size() - 1).getEndPos(), ruleMessage);
        String suggestion = resultTokens.get(resultTokens.size() - 2).getToken() + resultTokens.get(resultTokens.size() - 1).getToken();
        ruleMatch.addSuggestedReplacement(suggestion);
        setType(paraTokens.size() - 1, paraTokens.size() - 1, resultTokens.size() - 2, resultTokens.size() - 1, paraTokens, resultTokens, ruleMatch);
        tmpMatches.add(new AiRuleMatch(ruleMatch, resultTokens.get(j - 1).getStartPos(), resultTokens.get(resultTokens.size() - 1).getEndPos(),
            paraTokens.size() - 1, paraTokens.size() - 1, resultTokens.size() - 2, resultTokens.size() - 1));
      } else if (j < resultTokens.size() 
              && (!paraTokens.get(i - 1).getToken().equals(resultTokens.get(j - 1).getToken())
                  || (resultTokens.get(j).isNonWord() && !PUNCTUATION.matcher(resultTokens.get(j).getToken()).matches())
              && ((j + 2 > resultTokens.size()) 
                  || (!resultTokens.get(j + 1).getToken().equals("-") || !resultTokens.get(j + 2).getToken().equals(">")))
          )) {
        nRuleTokens++;
        if (nSentence > 0) {
          nSentence--;
        }
        int j1;
        for (j1 = j; j1 < resultTokens.size() && !OPENING_BRACKETS.matcher(resultTokens.get(j1).getToken()).matches(); j1++);
/*        
        String out = "";
        for (j1 = j; j1 < resultTokens.size() && !OPENING_BRACKETS.matcher(resultTokens.get(j1).getToken()).matches(); j1++) {
          out += resultTokens.get(j1).getToken() + "(" + j1 + ")";
        }
        WtMessageHandler.printToLogFile("j = " + j + ", j1 = " + j1 + ", r.length = " + resultTokens.size() + ", Result: " + out); 
*/
        if (j1 > resultTokens.size()) {
          j1 = resultTokens.size();
        }
//        WtMessageHandler.printToLogFile("j = " + j + ", j1 = " + j1 + ", r.length = " + resultTokens.size() + ", Result: " + out); 
        String suggestion = aiResultText.substring(resultTokens.get(j - 1).getStartPos(), resultTokens.get(j1 - 1).getEndPos());
//        WtMessageHandler.printToLogFile("StartPos = " + resultTokens.get(j - 1).getStartPos() 
//            + ", EndPos = " + resultTokens.get(j1 - 1).getEndPos() + ", suggestion = " + suggestion); 
        if (suggestion.isEmpty() || j != j1 || resultTokens.get(j - 1).isNonWord() 
            || linguServices.isCorrectSpell(suggestion, locale)) {
          if (!suggestion.equals(paraTokens.get(paraTokens.size() - 1).getToken())) {
            RuleMatch ruleMatch = new RuleMatch(this, null, paraTokens.get(paraTokens.size() - 1).getStartPos(), 
                paraTokens.get(paraTokens.size() - 1).getEndPos(), ruleMessage);
            ruleMatch.addSuggestedReplacement(suggestion);
            setType(paraTokens.size() - 1, paraTokens.size() - 1, j - 1, j1 - 1, paraTokens, resultTokens, ruleMatch);
            tmpMatches.add(new AiRuleMatch(ruleMatch, resultTokens.get(j - 1).getStartPos(), resultTokens.get(resultTokens.size() - 1).getEndPos(),
                paraTokens.size() - 1, paraTokens.size() - 1, j - 1, j1 - 1));
            if(debugMode > 1) {
              WtMessageHandler.printToLogFile("Text: " 
                  + paraText.substring(paraTokens.get(paraTokens.size() - 1).getStartPos(), paraTokens.get(paraTokens.size() - 1).getEndPos())
                  + "; suggestion: " + suggestion);
            }
          }
        }
      }
      if (tmpMatches.size() > 0) {
        boolean overSentenceEnd = j < resultTokens.size() && PUNCTUATION.matcher(resultTokens.get(j - 1).getToken()).matches();
        int nSenTokens = nSentence == 0 || sentenceEnds.size() == 1 ? sentenceEnds.get(0) : 
          sentenceEnds.get(sentenceEnds.size() - 1) - sentenceEnds.get(sentenceEnds.size() - 2);
        if (debugMode > 1) {
          WtMessageHandler.printToLogFile("AiDetectionRule: match: j < resultTokens.size(): mergeSentences: " + mergeSentences
              + ", nRuleTokens: " + nRuleTokens + ", nSenTokens: " + nSenTokens);
        }
        if (mergeSentences || overSentenceEnd || styleHintAssumed(nRuleTokens, nSenTokens, tmpMatches, paraTokens, resultTokens)) {
          if (showStylisticHints == 2) {
            int startPos;
            int endPos;
            int suggestionStart;
            int suggestionEnd;
            if (tmpMatches.size() > 0) {
              startPos = tmpMatches.get(0).ruleMatch.getFromPos();
              endPos = tmpMatches.get(tmpMatches.size() - 1).ruleMatch.getToPos();
              suggestionStart = tmpMatches.get(0).suggestionStart;
              suggestionEnd = tmpMatches.get(tmpMatches.size() - 1).suggestionEnd;
            } else {
              startPos = nSentence == 0 ? paraTokens.get(0).getStartPos() : paraTokens.get(sentenceEnds.get(nSentence - 1)).getStartPos();
              endPos = paraTokens.get(paraTokens.size() - 1).getEndPos();
              suggestionStart = resultTokens.get(lastResultStart).getStartPos();
              suggestionEnd = resultTokens.get(resultTokens.size() - 1).getEndPos();
            }
            RuleMatch ruleMatch = new RuleMatch(this, null, startPos, endPos, ruleMessage);
            String suggestion = aiResultText.substring(suggestionStart, suggestionEnd);
            ruleMatch.addSuggestedReplacement(suggestion);
            ruleMatch.setType(Type.Other);
            matches.add(ruleMatch);
            if (debugMode > 1) {
              WtMessageHandler.printToLogFile("AiDetectionRule: match: Stylistic hint: suggestion(2): " + suggestion);
            }
          }
        } else {
          addAllRuleMatches(matches, tmpMatches, resultTokens);
          if (debugMode > 1) {
            WtMessageHandler.printToLogFile("AiDetectionRule: match: add matches(2): " + tmpMatches.size()
            + ", total: " + matches.size());
          }
        }
      }
      if (debugMode > 1 && j < resultTokens.size()) {
        WtMessageHandler.printToLogFile("AiDetectionRule: match: j < resultTokens.size(): paraTokens.get(i - 1): " + paraTokens.get(i - 1).getToken()
            + ", resultTokens.get(j - 1): " + resultTokens.get(j - 1).getToken());
      }
      return toRuleMatchArray(matches);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
      return null;
    }
  }

  private boolean isCorrectSuggestion(String suggestion, WtAiToken singleWordToken) throws Throwable {
    return (suggestion.isEmpty() || singleWordToken == null || singleWordToken.isNonWord()
        || linguServices.isCorrectSpell(suggestion, locale));
  }

  private boolean isEndQoute(int j, List<WtAiToken> resultTokens) throws Throwable {
    return (j == resultTokens.size() - 1 && QUOTES.matcher(resultTokens.get(j).getToken()).matches());
  }

  private void setType(int nParaTokenStart, int nParaTokenEnd, int nResultTokenStart, int nResultTokenEnd,
      List<WtAiToken> paraTokens, List<WtAiToken> resultTokens, RuleMatch ruleMatch) throws Throwable {
    if (debugMode > 1) {
      WtMessageHandler.printToLogFile("nParaTokenStart: " + nParaTokenStart
          +", nParaTokenEnd: " + nParaTokenEnd
          + ", nResultTokenStart: " + nResultTokenStart
          +", nResultTokenEnd: " + nResultTokenEnd);
      WtMessageHandler.printToLogFile("ParaTokenStart: " + paraTokens.get(nParaTokenStart).getToken()
          +", ParaTokenEnd: " + paraTokens.get(nParaTokenEnd).getToken()
          + ", ResultTokenStart: " + resultTokens.get(nResultTokenStart).getToken()
          +", ResultTokenEnd: " + resultTokens.get(nResultTokenEnd).getToken());
      WtMessageHandler.printToLogFile("ParaTokenStart is Punctuation: " + PUNCTUATION.matcher(paraTokens.get(nParaTokenStart).getToken()).matches()
          + ", ResultTokenStart is Punctuation: " + PUNCTUATION.matcher(resultTokens.get(nResultTokenStart).getToken()).matches());
    }
    if (nParaTokenStart == nParaTokenEnd && nResultTokenStart == nResultTokenEnd
        && PUNCTUATION.matcher(paraTokens.get(nParaTokenStart).getToken()).matches()
        && PUNCTUATION.matcher(resultTokens.get(nResultTokenStart).getToken()).matches()) {
      //  other punctuation (e.g.: . --> ?
      ruleMatch.setMessage(ruleMessageOtherPunctuation);
      ruleMatch.setType(Type.Hint);
//      WtMessageHandler.printToLogFile("setType: ruleMessageOtherPunctuation set");
      return;
    }
    if (isNoneHintException(nParaTokenStart, nParaTokenEnd, nResultTokenStart, nResultTokenEnd,
        paraTokens, resultTokens, ruleMatch)) {
      //  NOTE: the type of message has to be set by the extended rule
      ruleMatch.setType(Type.Other);
      return;
    }
    if (isHintException(nParaTokenStart, nParaTokenEnd, nResultTokenStart, nResultTokenEnd, 
        paraTokens, resultTokens, ruleMatch)) {
      //  NOTE: the type of message has to be set by the extended rule
      ruleMatch.setType(Type.Hint);
      return;
    }
    if (nParaTokenStart == nParaTokenEnd && nResultTokenStart >= nResultTokenEnd
        && PUNCTUATION.matcher(paraTokens.get(nParaTokenStart).getToken()).matches()) {
      //  unnecessary punctuation
      ruleMatch.setMessage(ruleMessageUnnecessaryPunctuation);
      ruleMatch.setType(Type.Hint);
      return;
    }
    if (nParaTokenStart == nParaTokenEnd && nResultTokenStart == nResultTokenEnd) {
      if (!linguServices.isCorrectSpell(paraTokens.get(nParaTokenStart).getToken(), locale)) {
        ruleMatch.setMessage(ruleMessageMisspelling);
        ruleMatch.setType(Type.UnknownWord);
      } else if (shareLemma(paraTokens.get(nParaTokenStart), resultTokens.get(nResultTokenStart))) {
        ruleMatch.setMessage(ruleMessageGrammar);
        ruleMatch.setType(Type.Hint);
      } else if (isSimilarWord(paraTokens.get(nParaTokenStart).getToken(), resultTokens.get(nResultTokenStart).getToken())) {
        ruleMatch.setMessage(ruleMessageWordConfusion);
        ruleMatch.setType(Type.Hint);
      } else {
        ruleMatch.setType(Type.Other);
      } 
    } else if (nParaTokenStart + 1 == nParaTokenEnd && nResultTokenStart == nResultTokenEnd
        && ((PUNCTUATION.matcher(paraTokens.get(nParaTokenStart).getToken()).matches() 
              && paraTokens.get(nParaTokenEnd).getToken().equals(resultTokens.get(nResultTokenStart).getToken()))
            || (PUNCTUATION.matcher(paraTokens.get(nParaTokenEnd).getToken()).matches()
              && paraTokens.get(nParaTokenStart).getToken().equals(resultTokens.get(nResultTokenStart).getToken())))) {
      ruleMatch.setMessage(ruleMessageUnnecessaryPunctuation);
      ruleMatch.setType(Type.Hint);
    } else if (nParaTokenStart == nParaTokenEnd && nResultTokenStart + 1 == nResultTokenEnd
        && ((PUNCTUATION.matcher(resultTokens.get(nResultTokenStart).getToken()).matches() 
              && resultTokens.get(nResultTokenEnd).getToken().equals(paraTokens.get(nParaTokenStart).getToken()))
            || (PUNCTUATION.matcher(resultTokens.get(nResultTokenEnd).getToken()).matches()
              && resultTokens.get(nResultTokenStart).getToken().equals(paraTokens.get(nParaTokenStart).getToken())))) {
      ruleMatch.setMessage(ruleMessageMissingPunctuation);
      ruleMatch.setType(Type.Hint);
    } else {
      ruleMatch.setType(Type.Other);
    }
  }
  
  private boolean matchesShareLemma(int nParaTokenStart, int nParaTokenEnd, int nResultTokenStart, int nResultTokenEnd,
      List<WtAiToken> paraTokens, List<WtAiToken> resultTokens) throws Throwable {
    for (int i = nParaTokenStart; i <= nParaTokenEnd; i++) {
      for (int j = nResultTokenStart; j <= nResultTokenStart; j++) {
        if (shareLemma(paraTokens.get(i), resultTokens.get(j))) {
          return true;
        }
      }
    }
    return false;
  }
  
  private boolean styleHintAssumed(int nRuleTokens, int nSentTokens, 
      List<AiRuleMatch> aiMatches, List<WtAiToken> paraTokens, List<WtAiToken> resultTokens) throws Throwable {
    if (aiMatches.size() <= 2) {
      return false;
    }
    if (nRuleTokens >= nSentTokens / 2) {
      return true;
    }
    if (nSentTokens > 300 / STYLE_HINT_LIMIT && nRuleTokens >= 100 / STYLE_HINT_LIMIT ) {
      return true;
    }
    for (int i = 0; i < aiMatches.size(); i++) {
      for (int j = 0; j < aiMatches.size(); j++) {
        if (i != j) {
          if (matchesShareLemma(aiMatches.get(i).nParaTokenStart, aiMatches.get(i).nParaTokenEnd, 
              aiMatches.get(j).nResultTokenStart, aiMatches.get(j).nResultTokenEnd,
              paraTokens, resultTokens)) {
            return true;
          }
        }
      }
    }
    return false;
  }
  
  private boolean shareLemma(WtAiToken a, WtAiToken b) throws Throwable {
    for (AnalyzedToken t : a.getReadings()) {
      String lemma = t.getLemma();
      if (lemma != null && !lemma.isEmpty() && b.hasLemma(lemma)) {
        return true;
      }
    }
    return false;
  }
  
  private boolean mergeRuleMatchesOneTime(List<AiRuleMatch> aiMatches, List<WtAiToken> resultTokens) throws Throwable {
    for (int i = 0; i < aiMatches.size(); i++) {
      RuleMatch match1 = aiMatches.get(i).ruleMatch;
      if(match1.getSuggestedReplacements().get(0).isEmpty()) {
        String delTxt = paraText.substring(match1.getFromPos(), match1.getToPos()).trim();
        for (int j = 0; j < aiMatches.size(); j++) {
          RuleMatch match2 = aiMatches.get(j).ruleMatch;
          String suggestion = match2.getSuggestedReplacements().get(0);
          if (suggestion.contains(delTxt)) {
            if (i > j) {
              RuleMatch ruleMatch = new RuleMatch(this, match2.getSentence(), match2.getFromPos(), match1.getToPos(), ruleMessage);
              suggestion = aiResultText.substring(resultTokens.get(aiMatches.get(j).nResultTokenStart).getStartPos(), 
                  resultTokens.get(aiMatches.get(i).nResultTokenEnd).getEndPos());
              ruleMatch.addSuggestedReplacement(suggestion);
              ruleMatch.setType(Type.Other);
              aiMatches.add(i + 1, new AiRuleMatch(ruleMatch, 
                  aiMatches.get(j).suggestionStart, aiMatches.get(i).suggestionEnd,
                  aiMatches.get(j).nParaTokenStart, aiMatches.get(i).nParaTokenEnd, 
                  aiMatches.get(j).nResultTokenStart, aiMatches.get(i).nResultTokenEnd));
              for (int k = i; k >= j; k--) {
                aiMatches.remove(k);
              }
              return true;
            } else if (j > i) {
              boolean startsWithWhiteChar = paraText.substring(match1.getFromPos(), match1.getFromPos() + 1).trim().isEmpty();
              int fromPos = startsWithWhiteChar ? match1.getFromPos() + 1 : match1.getFromPos();
              RuleMatch ruleMatch = new RuleMatch(this, match1.getSentence(), fromPos, match2.getToPos(), ruleMessage);
              suggestion = aiResultText.substring(resultTokens.get(aiMatches.get(i).nResultTokenStart).getStartPos(), 
                  resultTokens.get(aiMatches.get(j).nResultTokenEnd).getEndPos());
              ruleMatch.addSuggestedReplacement(suggestion);
              ruleMatch.setType(Type.Other);
              aiMatches.add(j + 1, new AiRuleMatch(ruleMatch, 
                  aiMatches.get(i).suggestionStart, aiMatches.get(j).suggestionEnd,
                  aiMatches.get(i).nParaTokenStart, aiMatches.get(j).nParaTokenEnd, 
                  aiMatches.get(i).nResultTokenStart, aiMatches.get(j).nResultTokenEnd));
              for (int k = j; k >= i; k--) {
                aiMatches.remove(k);
              }
              return true;
            }
          }
        }
      }
    }
    return false;
  }
 
  private void mergeRuleMatches(List<AiRuleMatch> aiMatches, List<WtAiToken> resultTokens) throws Throwable {
    while (mergeRuleMatchesOneTime(aiMatches, resultTokens));
  }
  
  private void addAllRuleMatches(List<RuleMatch> matches, List<AiRuleMatch> aiMatches, List<WtAiToken> resultTokens) throws Throwable {
    mergeRuleMatches(aiMatches, resultTokens);
    for (AiRuleMatch match : aiMatches) {
      if (showStylisticHints > 0 || match.ruleMatch.getType() == Type.Hint) {
        matches.add(match.ruleMatch);
      }
    }
  }
  
  private boolean hasAllChars (String sFirst, String sSecond, int diffChars) throws Throwable {
    List<Character> cList = new ArrayList<>();
    for (char c : sFirst.toCharArray()) {
        cList.add(c);
    }
    for (char c : sSecond.toCharArray()) {
      int indx = cList.indexOf(c);
      if (indx >= 0) {
        cList.remove(indx);
      } else {
        if (diffChars <= 0) {
          return false;
        } else {
          diffChars--;
        }
      }
    }
    return true;
  }
  
  private boolean isSimilarWord (String first, String second) throws Throwable {
    if (first.length() < 2 || second.length() < 2) {
      return false;
    }
    if (first.length() == second.length()) {
      return hasAllChars (first, second, 1);
    } else if (first.length() == second.length() + 1) {
      return hasAllChars (first, second, 0);
    } else if (first.length() + 1 == second.length()) {
      return hasAllChars (second, first, 0);
    }
    return false;
  }
  
  /**
   * Set language specific exceptions to set Color for specific Languages
   */
  public boolean isNoneHintException(int nParaStart, int nParaEnd, int nResultStart, int nResultEnd, 
      List<WtAiToken> paraTokens, List<WtAiToken> resultTokens, RuleMatch ruleMatch) throws Throwable {
    return false;   
  }

  /**
   * Set language specific exceptions to set Color for specific Languages
   */
  public boolean isHintException(int nParaStart, int nParaEnd, int nResultStart, int nResultEnd, 
      List<WtAiToken> paraTokens, List<WtAiToken> resultTokens, RuleMatch ruleMatch) throws Throwable {
    return false;   
  }

  /**
   * Set language specific exceptions to handle change as a match
   */
  public boolean isMatchException(int nParaStart, int nParaEnd, 
      int nResultStart, int nResultEnd, List<WtAiToken> paraTokens, List<WtAiToken> resultTokens) throws Throwable {
    return false;   
  }

  @Override
  public int minToCheckParagraph() {
    return 0;
  }

  @Override
  public String getId() {
    return WtOfficeTools.AI_GRAMMAR_RULE_ID;
  }

  @Override
  public String getDescription() {
    return messages.getString("loAiRuleDescription");
  }
  
  class AiRuleMatch {
    public final RuleMatch ruleMatch;
    public final int suggestionStart;
    public final int suggestionEnd;
    public final int nParaTokenStart;
    public final int nParaTokenEnd;
    public final int nResultTokenStart;
    public final int nResultTokenEnd;
    
    AiRuleMatch(RuleMatch ruleMatch, int suggestionStart, int suggestionEnd, 
        int nParaTokenStart, int nParaTokenEnd, int nResultTokenStart, int nResultTokenEnd) {
      this.ruleMatch = ruleMatch;
      this.suggestionStart = suggestionStart;
      this.suggestionEnd = suggestionEnd;
      this.nParaTokenStart = nParaTokenStart;
      this.nParaTokenEnd = nParaTokenEnd;
      this.nResultTokenStart = nResultTokenStart;
      this.nResultTokenEnd = nResultTokenEnd;
    }
  }

}
