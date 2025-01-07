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
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.GZIPOutputStream;

import org.json.JSONObject;
import org.languagetool.AnalyzedSentence;
import org.languagetool.AnalyzedToken;
import org.languagetool.AnalyzedTokenReadings;
import org.languagetool.JLanguageTool;
import org.languagetool.Language;
import org.languagetool.Languages;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;

import com.sun.star.lang.Locale;

/**
 * Class to cache analyzed Sentences
 * For test reasons only / not shown to the user
 * @since 1.0
 * @author Fred Kruse
 */
public class WtAnalyzedParagraphsCache {
  private JSONObject doc = new JSONObject();
  private String locale;
  
  WtAnalyzedParagraphsCache(WtDocumentsHandler mDocHandler) {
    WtSingleDocument document = mDocHandler.getCurrentDocument();
    if (document == null) {
      return;
    }
    WtDocumentCache docCache = document.getDocumentCache();
    if (docCache == null) {
      return;
    }
    Locale tmpLocale = docCache.getFlatParagraphLocale(0);
    if (tmpLocale == null || !WtDocumentsHandler.hasLocale(tmpLocale)) {
      locale = null;
      return;
    }
    locale = WtOfficeTools.localeToString(tmpLocale);
    Language language = Languages.getLanguageForShortCode(WtOfficeTools.localeToString(tmpLocale));
    JLanguageTool lt = new JLanguageTool(language);
    try {
      TextParagraph tPara = new TextParagraph(WtDocumentCache.CURSOR_TYPE_TEXT, 0);
      List<String> jParagraphs = new ArrayList<String>();
      for (int n = 0; n < docCache.textSize(tPara); n++) {
        tPara = new TextParagraph(WtDocumentCache.CURSOR_TYPE_TEXT, n);
        String para = docCache.getTextParagraph(tPara);
        List<String> sentences = lt.sentenceTokenize(para);
        List<String> jSentences = new ArrayList<String>();
        JSONObject jParagraph = new JSONObject();
        for (String sentence : sentences) {
          JAnalyzedSentence jSentence = new JAnalyzedSentence(lt.getAnalyzedSentence(sentence));
          jSentences.add(jSentence.getJSON());
        }
        jParagraph.put("pa", jSentences);
        jParagraphs.add(jParagraph.toString());
      }
      doc.put("locale", locale);
      doc.put("paragraphs", jParagraphs);
      writeIntoFile();
    } catch (IOException e) {
      WtMessageHandler.printException(e);
      return;
    }
  }
  
  String getLocaleAsString() {
    return locale;
  }
  
  void writeIntoFile() {
    try {
      File tmpDir = new File(WtOfficeTools.getWtConfigDir(), "tmp");
      if (tmpDir != null && !tmpDir.exists()) {
        tmpDir.mkdirs();
      }
      File tmpCacheFile = new File(tmpDir, "tmp_AnalyzedParagraphsCache");
      if (tmpCacheFile.exists()) {
        tmpCacheFile.delete();
      }
      GZIPOutputStream fileOut = new GZIPOutputStream(new FileOutputStream(tmpCacheFile.getAbsolutePath()));
      OutputStreamWriter out = new OutputStreamWriter(fileOut, StandardCharsets.UTF_8);
      out.write(doc.toString());
      out.close();
      fileOut.close();
      File cacheFile = new File(tmpDir, "AnalyzedParagraphsCache");
      if (cacheFile.exists()) {
        cacheFile.delete();
      }
      tmpCacheFile.renameTo(cacheFile);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
    }
  }
  
  private class JAnalyzedSentence {
    private List<JAnalyzedTokenReadings> tokens = new ArrayList<JAnalyzedTokenReadings>();
    
    JAnalyzedSentence(AnalyzedSentence sentences) {
      for (AnalyzedTokenReadings token : sentences.getTokensWithoutWhitespace()) {
        tokens.add(new JAnalyzedTokenReadings(token));
      }
    }

    String getJSON() {
      JSONObject obj = new JSONObject();
      List<String> jSentences = new ArrayList<String>();
      for (JAnalyzedTokenReadings token : tokens) {
        jSentences.add(token.getJSON());
      }
      obj.put("aS", jSentences);
      return obj.toString();
    }
    
  }
  
  private class JAnalyzedTokenReadings {
    private List<JAnalyzedToken> tokenReadings = new ArrayList<JAnalyzedToken>();
    private String token;
    private int startPos;
    
    JAnalyzedTokenReadings(AnalyzedTokenReadings anTokReadings) {
      token = anTokReadings.getToken();
      startPos = anTokReadings.getStartPos();
      for (AnalyzedToken aToken : anTokReadings.getReadings()) {
        tokenReadings.add(new JAnalyzedToken(aToken));
      }
    }
    
    String getJSON() {
      JSONObject obj = new JSONObject();
      obj.put("to", token);
      obj.put("sP", startPos);
      List<String> jReadings = new ArrayList<String>();
      for (JAnalyzedToken reading : tokenReadings) {
        jReadings.add(reading.getJSON());
      }
      obj.put("tR", jReadings);
      return obj.toString();
    }
    
  }
  
  private class JAnalyzedToken {
    private String posTag;
    private String lemma;
    private boolean hasNoTag = false;
    
    JAnalyzedToken(AnalyzedToken token) {
      posTag =  token.getPOSTag() == null ? "" : token.getPOSTag();
      lemma = token.getLemma() == null ? "" : token.getLemma();
      hasNoTag = token.hasNoTag();
    }
    
    String getJSON() {
      JSONObject obj = new JSONObject();
      obj.put("pT", posTag);
      obj.put("le", lemma);
      obj.put("hT", hasNoTag);
      return obj.toString();
    }
    
  }

}

