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

import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;

import org.languagetool.AnalyzedSentence;
import org.languagetool.AnalyzedTokenReadings;
import org.writingtool.WtDocumentCache;
import org.writingtool.WtLanguageTool;
import org.writingtool.WtProofreadingError;
import org.writingtool.WtSingleCheck;
import org.writingtool.WtSingleDocument;
import org.writingtool.WtSuggestionStore;
import org.writingtool.WtDocumentCache.AnalysedText;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.WtDocumentsHandler.WaitDialogThread;
import org.writingtool.aisupport.WtAiRemote.AiCommand;
import org.writingtool.config.WtConfiguration;
import org.writingtool.dialogs.WtAiDialog;
import org.writingtool.dialogs.WtAiResultDialog;
import org.writingtool.dialogs.WtOptionPane;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtViewCursorTools;

import com.sun.star.lang.Locale;
import com.sun.star.lang.XComponent;

/**
 * Class to execute changes of paragraphs by AI
 * @since 1.0
 * @author Fred Kruse
 */
public class WtAiParagraphChanging extends Thread {

  public final static int MAX_WORDS = 1000;

  private static final ResourceBundle messages = WtOfficeTools.getMessageBundle();

  private int debugMode = WtOfficeTools.DEBUG_MODE_AI;   //  should be 0 except for testing

  public final static String WAIT_TITLE = messages.getString("loAiWaitDialogTitle");
  public final static String WAIT_MESSAGE = messages.getString("loAiWaitDialogMessage");

  public final static WtSuggestionStore lastWords = new WtSuggestionStore(MAX_WORDS);
  
  private final WtSingleDocument document;
  private final WtConfiguration config;
  private final AiCommand commandId;
  
  private static WtAiResultDialog resultDialog = null;
  private static WtAiDialog aiDialog = null;
  private WaitDialogThread waitDialog = null;
  private String word;
  
  public WtAiParagraphChanging(WtSingleDocument document, WtConfiguration config, AiCommand commandId) {
    this.document = document;
    this.config = config;
    this.commandId = commandId;
  }
  
  @Override
  public void run() {
    runAiChangeOnParagraph();
  }
  
  public static WtAiDialog getAiDialog() {
    return aiDialog;
  }
  
  public static void setCloseAiDialog() {
    aiDialog = null;
  }
  
  public static WtAiResultDialog getAiResultDialog() {
    return resultDialog;
  }
  
  public static void setCloseAiResultDialog() {
    resultDialog = null;
  }
  
  private void runAiChangeOnParagraph() {
    try {
      if (commandId == AiCommand.GeneralAi) {
        if (aiDialog == null) {
          waitDialog = new WaitDialogThread(WAIT_TITLE, WAIT_MESSAGE);
          aiDialog = new WtAiDialog(document, waitDialog, messages);
          aiDialog.start();
        } else {
          aiDialog.toFront();
        }
        return;
      } else if (commandId == AiCommand.SynonymsOfWord) {
        showSynonyms();
        return;
      }
      if (debugMode > 1) {
        WtMessageHandler.printToLogFile("AiParagraphChanging: runAiChangeOnParagraph: commandId: " + commandId);
      }
      XComponent xComponent = document.getXComponent();
      WtViewCursorTools viewCursor = new WtViewCursorTools(xComponent);
      TextParagraph tPara = viewCursor.getViewCursorParagraph();
      viewCursor.selectParagraphFromViewCursor();
      WtDocumentCache docCache = document.getDocumentCache();
      Locale locale = docCache.getTextParagraphLocale(tPara);
      String text = commandId == AiCommand.SynonymsOfWord ? viewCursor.getWordFromViewCursor() : docCache.getTextParagraph(tPara);
      if (text == null || text.trim().isEmpty()) {
        return;
      }
      String instruction = null;
      boolean onlyPara = false;
      float temp = WtAiRemote.CORRECT_TEMPERATURE;
      String labelText;
      if (commandId == AiCommand.CorrectGrammar) {
        instruction = WtAiRemote.getInstruction(WtAiRemote.CORRECT_INSTRUCTION, locale);
        onlyPara = true;
        labelText = "";
      } else if (commandId == AiCommand.ImproveStyle) {
        instruction = WtAiRemote.getInstruction(WtAiRemote.STYLE_INSTRUCTION, locale);
        onlyPara = true;
        labelText = messages.getString("loMenuAiStyleCommand") + ":";
      } else if (commandId == AiCommand.ReformulateText) {
        instruction = WtAiRemote.getInstruction(WtAiRemote.REFORMULATE_INSTRUCTION, locale);
        onlyPara = true;
        temp = WtAiRemote.REFORMULATE_TEMPERATURE;
        labelText = messages.getString("loMenuAiReformulateCommand") + ":";
      } else if (commandId == AiCommand.ExpandText) {
        instruction = WtAiRemote.getInstruction(WtAiRemote.EXPAND_INSTRUCTION, locale);
        temp = WtAiRemote.EXPAND_TEMPERATURE;
        labelText = messages.getString("loMenuAiExpandCommand") + ":";
      } else {
        instruction = WtOptionPane.showInputDialog(null, messages.getString("loMenuAiGeneralCommandMessage"), 
          messages.getString("loMenuAiGeneralCommandTitle"), WtOptionPane.QUESTION_MESSAGE);
        temp = WtAiRemote.EXPAND_TEMPERATURE;
        labelText = messages.getString("loMenuAiGeneralCommand") + ":";
      }
//      WtMessageHandler.printToLogFile("AiParagraphChanging: runAiChangeOnParagraph: Result Dialog " + (resultDialog != null ? "!=" : "==") + " null");
      if (resultDialog != null) {
//        WtMessageHandler.printToLogFile("AiParagraphChanging: runAiChangeOnParagraph: Result Dialog closed");
        resultDialog.closeDialog();
      }
      waitDialog = new WaitDialogThread(WAIT_TITLE, WAIT_MESSAGE);
      waitDialog.start();
      WtAiRemote aiRemote = new WtAiRemote(document.getMultiDocumentsHandler(), config);
      if (debugMode > 1) {
        WtMessageHandler.printToLogFile("AiParagraphChanging: runAiChangeOnParagraph: instruction: " + instruction + ", text: " + text);
      }
      String output = aiRemote.runInstruction(instruction, text, temp, 1, locale, onlyPara);
      if (debugMode > 1) {
        WtMessageHandler.printToLogFile("AiParagraphChanging: runAiChangeOnParagraph: output: " + output);
      }
      resultDialog = new WtAiResultDialog(document, messages, labelText, false);
      if (output == null) {
        output = "";
      }
      resultDialog.setResult(output, tPara);
      resultDialog.start();
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    }
    if (waitDialog != null) {
      waitDialog.close();
      waitDialog = null;
    }
  }
  
  /** 
   * Inserts a Text to cursor position
   */
  public static void insertText(String text, XComponent xComponent, TextParagraph yPara, boolean override) {
    WtViewCursorTools vCursor = new WtViewCursorTools(xComponent);
    vCursor.setTextViewCursor(0, yPara);
    vCursor.insertText(text, override);
  }

  public static void insertText(String text, XComponent xComponent, boolean override) {
    WtViewCursorTools vCursor = new WtViewCursorTools(xComponent);
    vCursor.insertText(text, override);
  }

  /**
   * change word in a paragraph
   */
  public static void changeWordInParagraph(TextParagraph tPara, int nStart, int nLength, String replace, 
      WtSingleDocument document) throws Throwable {
    WtDocumentCache docCache = document.getDocumentCache();
    int nFPara = docCache.getFlatParagraphNumber(tPara);
    String sPara = docCache.getFlatParagraph(nFPara);
    String sEnd = (nStart + nLength < sPara.length() ? sPara.substring(nStart + nLength) : "");
    sPara = sPara.substring(0, nStart) + replace + sEnd;
    document.getFlatParagraphTools().changeTextOfParagraph(nFPara, nStart, nLength, replace);
    docCache.setFlatParagraph(nFPara, sPara);
    document.removeResultCache(nFPara, true);
    document.removeIgnoredMatch(nFPara, true);
    document.removePermanentIgnoredMatch(nFPara, true);
  }

  public void showSynonyms() {
    String[] synonyms = null;
    TextParagraph tPara = null;
    int nStart = 0;
    int nLength = 0;
    try {
      if (resultDialog != null) {
        resultDialog.closeDialog();
      }
      waitDialog = new WaitDialogThread(WAIT_TITLE, WAIT_MESSAGE);
      waitDialog.start();
      XComponent xComponent = document.getXComponent();
      WtViewCursorTools viewCursor = new WtViewCursorTools(xComponent);
      tPara = viewCursor.getViewCursorParagraph();
      int nChar = viewCursor.getViewCursorCharacter();
      viewCursor.selectWordFromViewCursor();
      WtDocumentCache docCache = document.getDocumentCache();
      WtLanguageTool lt = document.getMultiDocumentsHandler().getLanguageTool();
      if (lt != null) {
        int nFPara = docCache.getFlatParagraphNumber(tPara);
        WtProofreadingError error = getProofreadingError(nChar, nFPara, docCache, lt, viewCursor);
        if (error == null) {
          if (debugMode > 1) {
            WtMessageHandler.printToLogFile("showSynonyms: error == null, nFPara: " + nFPara + ", nChar: " + nChar);
          }
          synonyms = new String[0];
        } else {
          nStart = error.nErrorStart;
          nLength = error.nErrorLength;
          synonyms = error.aSuggestions;
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
    }
    String labelText = messages.getString("loMenuAiSynnomsOfWordCommand") + " '" + word + "':";
    resultDialog = new WtAiResultDialog(document, messages, labelText, true);
    if (synonyms == null) {
      if (debugMode > 1) {
        WtMessageHandler.printToLogFile("showSynonyms: synonyms == null");
      }
      synonyms = new String[0];
    }
    resultDialog.setResultList(synonyms, tPara, nStart, nLength);
    if (waitDialog != null) {
      waitDialog.close();
      waitDialog = null;
    }
    resultDialog.start();
  }
  
  public WtProofreadingError getProofreadingError(int nChar, int nFPara, 
      WtDocumentCache docCache, WtLanguageTool lt, WtViewCursorTools viewCursor) throws Throwable {
    return getProofreadingError(nChar, null, nFPara, docCache, lt, viewCursor);
  }

  public WtProofreadingError getProofreadingError(WtProofreadingError error, int nFPara, 
      WtDocumentCache docCache, WtLanguageTool lt) throws Throwable {
    return getProofreadingError(0, error, nFPara, docCache, lt, null);
  }

  public WtProofreadingError getProofreadingError(int nChar, WtProofreadingError error, int nFPara, 
      WtDocumentCache docCache, WtLanguageTool lt, WtViewCursorTools viewCursor) throws Throwable {
    if (error != null) {
      nChar = error.nErrorStart;
    }
    AnalysedText analysedText = docCache.getOrCreateAnalyzedParagraph(nFPara, lt);
    if (analysedText == null) {
      return null;
    }
    List<AnalyzedSentence> analyzedSentences = analysedText.analyzedSentences;
    int nPos = 0;
    for (AnalyzedSentence analyzedSentence : analyzedSentences) {
      for (AnalyzedTokenReadings token : analyzedSentence.getTokens()) {
        if (nChar >= token.getStartPos() + nPos && nChar < token.getEndPos() + nPos) {
          if (token.isPosTagUnknown() || token.isNonWord()) {
            WtMessageHandler.printToLogFile("getProofreadingError: token isPosTagUnknown || isNonWord: " + token.getToken());
            return null;
          }
          String[] synonyms  = getListOfSynonyms(token, nFPara, docCache);
          word = token.getToken();
          if (synonyms == null) {
            synonyms = new String[0];
          }
          if (error == null) {
            error = new WtProofreadingError();
            error.nErrorLength = token.getEndPos() - token.getStartPos();
            error.nErrorStart = token.getStartPos() + nPos;
            error = WtSingleCheck.correctRuleMatchWithFootnotes(error, 
                docCache.getFlatParagraphFootnotes(nFPara), docCache.getFlatParagraphDeletedCharacters(nFPara),
                nFPara, docCache.getHiddenCharactersMap());
          }
          error.aSuggestions = synonyms;
          return error;
        }
      }
      nPos += analyzedSentence.getCorrectedTextLength();
    }
    WtMessageHandler.printToLogFile("getProofreadingError: token not found: nChar: " + nChar + ", nFPara: " + nFPara);
    return null;
  }
  
  public String[] getListOfSynonyms(AnalyzedTokenReadings token, int nFPara, WtDocumentCache docCache) throws Throwable {
    Locale locale = docCache.getFlatParagraphLocale(nFPara);
    return getListOfSynonyms(token.getToken(), locale);
  }
  
  private void addSynonym(String synonym, String word, List<String> synnonymList) {
    if (synonym != null) {
      synonym = synonym.trim();
      if (!synonym.isEmpty() && !word.equals(synonym) && !synnonymList.contains(synonym)) {
        synnonymList.add(synonym);
      }
    }
  }
  
  private String[] getListOfSynonyms(String word, Locale locale) throws Throwable {
    WtMessageHandler.printToLogFile("getListOfSynonyms: word: " + word);
    String[] synonymArray = lastWords.getSuggestions(word, locale);
    if (synonymArray != null) {
      return synonymArray;
    } else if (debugMode > 1) {
      WtMessageHandler.printToLogFile("getListOfSynonyms: Synnonym Array == null: word: " + word + ", locale: " 
            + WtOfficeTools.localeToString(locale));
    }
    WtAiRemote aiRemote = new WtAiRemote(document.getMultiDocumentsHandler(), config);
    String instruction = WtAiRemote.getInstruction(WtAiRemote.SYNONYMS_INSTRUCTION, locale);
    String output = aiRemote.runInstruction(instruction, word, WtAiRemote.SYNONYM_TEMPERATURE, 0, locale, false);
    List<String> synonyms = new ArrayList<>();
    if (debugMode > 1) {
      WtMessageHandler.printToLogFile("getListOfSynonyms: output: " + output);
    }
    if (output == null || output.isBlank()) {
      if (debugMode > 0) {
        WtMessageHandler.printToLogFile("getListOfSynonyms: output == null");
      }
      return new String[0];
    }
    String[] outp = output.split(":");
    if (outp.length == 2) {
      output = outp[1];
    }
    outp = output.split("[\n\r]");
    for (String out : outp) {
      String[] otp = out.split("[,;]");
      for (String ot : otp) {
        String[] o = ot.split("\\.");
        if (o.length == 1) {
          o = ot.split("-");
          if (o.length == 1) {
            addSynonym(o[0], word, synonyms);
          } else if(o.length == 2) {
            addSynonym(o[1], word, synonyms);
          }
        } else if(o.length == 2) {
          addSynonym(o[1], word, synonyms);
        }
      }
    }
    synonymArray = synonyms.toArray(new String[0]);
    if (debugMode > 1) {
      WtMessageHandler.printToLogFile("getListOfSynonyms: Add Synnonym Array: word: " + word + ", locale: " 
        + WtOfficeTools.localeToString(locale) + ", synonyms size: " + synonymArray.length);
    }
    lastWords.addSuggestions(word, locale, synonymArray);
    return synonymArray;
  }

}
