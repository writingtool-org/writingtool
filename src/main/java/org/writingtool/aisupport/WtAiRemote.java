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
package org.writingtool.aisupport;

import java.io.BufferedReader;
import java.io.DataOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.ConnectException;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.util.ResourceBundle;
import java.util.Set;

import org.json.JSONArray;
import org.json.JSONObject;
import org.writingtool.WtDocumentsHandler;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;

import com.sun.star.lang.Locale;


/**
 * Class to communicate with a AI API
 * @since 6.5
 * @author Fred Kruse
 */
public class WtAiRemote {
  
  private final static int REMOTE_TRIALS = 5;
  private final static int BUFFER_SIZE = 20000;
  
  private static final ResourceBundle messages = WtOfficeTools.getMessageBundle();
  public final static String CORRECT_INSTRUCTION = "Output the grammatically and orthographically corrected text";
  public final static String STYLE_INSTRUCTION = "Output the stylistically improved text";
  public final static String REFORMULATE_INSTRUCTION = "Output the reformulated text";
  public final static String EXPAND_INSTRUCTION = "Output the expanded text";
  
  public static enum AiCommand { CorrectGrammar, ImproveStyle, ReformulateText, ExpandText, GeneralAi };
  
  private static boolean isRunning = false;

  private enum AiType { EDITS, COMPLETIONS, CHAT }
  
  boolean debugModeTm = WtOfficeTools.DEBUG_MODE_TM;
  boolean debugMode = WtOfficeTools.DEBUG_MODE_AI;
  
  private final WtDocumentsHandler documents;
  private final WtConfiguration config;
  private final String apiKey;
  private final String model;
  private final String url;
  private final String imgApiKey;
  private final String imgModel;
  private final String imgUrl;
  private final String ttsApiKey;
  private final String ttsModel;
  private final String ttsUrl;
  private final AiType aiType;
  
  public WtAiRemote(WtDocumentsHandler documents, WtConfiguration config) throws Throwable {
    this.documents = documents;
    this.config = config;
    apiKey = config.aiApiKey();
    model = config.aiModel();
    url = config.aiUrl();
    imgApiKey = config.aiImgApiKey();
    imgModel = config.aiImgModel();
    imgUrl = config.aiImgUrl();
    ttsApiKey = config.aiTtsApiKey();
    ttsModel = config.aiTtsModel();
    ttsUrl = config.aiTtsUrl();
    if (url.endsWith("/edits/") || url.endsWith("/edits")) {
      aiType = AiType.EDITS;
    } else if (url.endsWith("/chat/completions/") || url.endsWith("/chat/completions")) {
      aiType = AiType.CHAT;
    } else if (url.endsWith("/completions/") || url.endsWith("/completions")) {
      aiType = AiType.COMPLETIONS;
    } else {
      aiType = AiType.CHAT;
    }
  }

  public String runInstruction(String instruction, String text, float temperature, 
      int seed, Locale locale, boolean onlyOneParagraph) throws Throwable {
    try {
      while (isRunning) {
        try {
          Thread.sleep(100);
        } catch (InterruptedException e) {
          WtMessageHandler.printException(e);
        }
      }
      isRunning = true;
      return runInstruction_intern(instruction, text, temperature, seed, locale, onlyOneParagraph);
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      return null;
    } finally {
      isRunning = false;
    }
  }
  
  public String runImgInstruction(String instruction, String exclude, int step, int seed, int size) throws Throwable {
    try {
      while (isRunning) {
        try {
          Thread.sleep(100);
        } catch (InterruptedException e) {
          WtMessageHandler.printException(e);
        }
      }
      isRunning = true;
      return runImgInstruction_intern(instruction, exclude, step, seed, size);
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      return null;
    } finally {
      isRunning = false;
    }
  }

  public String runTtsInstruction(String text, String filename) throws Throwable {
    try {
      while (isRunning) {
        try {
          Thread.sleep(100);
        } catch (InterruptedException e) {
          WtMessageHandler.printException(e);
        }
      }
      isRunning = true;
      return runTtsInstruction_intern(text, filename);
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      return null;
    } finally {
      isRunning = false;
    }
  }

  
  private String runInstruction_intern(String instruction, String text, float temperature, 
      int seed, Locale locale, boolean onlyOneParagraph) throws Throwable {
    if (instruction == null || text == null) {
      return null;
    }
    instruction = instruction.trim();
    text = text.trim();
    if (onlyOneParagraph && (instruction.isEmpty() || text.isEmpty())) {
      return "";
    }
    if (instruction.isEmpty()) {
      if (text.isEmpty()) {
        return text;
      }
      instruction = text;
      text = null;
    } else if (text.isEmpty()) {
      text = null;
    }
    long startTime = 0;
    if (debugModeTm) {
      startTime = System.currentTimeMillis();
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("AiRemote: runInstruction: Ask AI started! URL: " + url);
    }
//    String langName = getLanguageName(locale);
    String langName = locale.Language;
    String org = text == null ? instruction : text;
    if (text != null) {
      text = text.replace("\n", "\r").replace("\r", " ").replace("\"", "\\\"").replace("\t", " ");
    }
    String urlParameters;
//    instruction = addLanguageName(instruction, locale);
    if (aiType == AiType.CHAT) {
      urlParameters = "{\"model\": \"" + model + "\", " 
//          + "\"response_format\": { \"type\": \"json_object\" }, "
          + "\"language\": \"" + langName + "\", "
          + "\"messages\": [ { \"role\": \"user\", "
          + "\"content\": \"" + instruction + (text == null ? "" : ": {" + text + "}") + "\" } ], "
          + (seed > 0 ? "\"seed\": " + seed + ", " : "")
          + "\"temperature\": " + temperature + "}";
    } else if (aiType == AiType.EDITS) {
      urlParameters = "{\"model\": \"" + model + "\", " 
//          + "\"response_format\": { \"type\": \"json_object\" },"
          + "\"instruction\": \"" + instruction + "\","
          + "\"input\": \"" + text + "\", "
//          + "\"seed\": 1, "
          + "\"temperature\": " + temperature + "}";
    } else {
      urlParameters = "{\"model\": \"" + model + "\", " 
//          + "\"response_format\": { \"type\": \"json_object\" },"
          + "\"prompt\": \"" + instruction + ": {" + text + "}\", "
//          + "\"seed\": 1, "
          + "\"temperature\": " + temperature + "}";
    }

    byte[] postData = urlParameters.getBytes(StandardCharsets.UTF_8);
    
    URL checkUrl;
    try {
      checkUrl = new URL(url);
    } catch (MalformedURLException e) {
      WtMessageHandler.showError(e);
      stopAiRemote();
      return null;
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("AiRemote: runInstruction: postData: " + urlParameters);
    }
    HttpURLConnection conn = null;
    try {
      conn = getConnection(postData, checkUrl, apiKey);
    } catch (RuntimeException e) {
      WtMessageHandler.printException(e);
      stopAiRemote();
      return null;
    }
    int trials = 0;
    while (trials < REMOTE_TRIALS) {
      trials++;
      try {
        if (conn.getResponseCode() == HttpURLConnection.HTTP_OK) {
          try (InputStream inputStream = conn.getInputStream()) {
            String out = readStream(inputStream, "utf-8");
            out = parseJasonOutput(out);
            if (out == null) {
              return null;
            }
            out = filterOutput (out, org, instruction, onlyOneParagraph);
            if (debugModeTm) {
              long runTime = System.currentTimeMillis() - startTime;
              WtMessageHandler.printToLogFile("AiRemote: runInstruction: Time to generate Answer: " + runTime);
            }
            return out;
          }
        } else {
          try (InputStream inputStream = conn.getErrorStream()) {
            String error = readStream(inputStream, "utf-8");
            WtMessageHandler.printToLogFile("Got error: " + error + " - HTTP response code " + conn.getResponseCode());
            WtMessageHandler.printToLogFile("urlParameters: " + urlParameters);
            stopAiRemote();
            return null;
          }
        }
      } catch (ConnectException e) {
        if (trials >= REMOTE_TRIALS) {
          WtMessageHandler.printToLogFile("Could not connect to server at: " + url);
          WtMessageHandler.printException(e);
          stopAiRemote();
        }
      } catch (Exception e) {
        if (trials >= REMOTE_TRIALS) {
          WtMessageHandler.showError(e);
          stopAiRemote();
        }
      } finally {
        conn.disconnect();
      }
    }
    return null;
  }
  
  private String runImgInstruction_intern(String instruction, String exclude, int step, int seed, int size) throws Throwable {
    if (instruction == null || exclude == null) {
      return null;
    }
    instruction = instruction.trim();
    if (instruction.isEmpty()) {
      return "";
    }
    exclude = exclude.trim();
    if (!exclude.isEmpty()) {
      instruction += "|" + exclude;
    }
    if (size != 128 && size != 256 && size != 512) {
      size = 256;
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("AiRemote: runImgInstruction: Ask AI started! URL: " + url);
    }
    String urlParameters = "{\"model\": \"" + imgModel + "\", " 
        + "\"prompt\": \"" + instruction + "\", "
        + (seed > 0 ? "\"seed\": " + seed + ", " : "")
        + "\"size\": \"" + size + "x" + size + "\", "
        + "\"step\": " + step + "}";
    
    byte[] postData = urlParameters.getBytes(StandardCharsets.UTF_8);
    
    URL checkUrl;
    try {
      checkUrl = new URL(imgUrl);
    } catch (MalformedURLException e) {
      WtMessageHandler.showError(e);
      stopAiImgRemote();
      return null;
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("AiRemote: runImgInstruction: postData: " + urlParameters);
    }
    HttpURLConnection conn = getConnection(postData, checkUrl, imgApiKey);
    try {
      if (conn.getResponseCode() == HttpURLConnection.HTTP_OK) {
        try (InputStream inputStream = conn.getInputStream()) {
          String out = readStream(inputStream, "utf-8");
          return parseJasonImgOutput(out);
        }
      } else {
        try (InputStream inputStream = conn.getErrorStream()) {
          String error = readStream(inputStream, "utf-8");
          WtMessageHandler.printToLogFile("Got error: " + error + " - HTTP response code " + conn.getResponseCode());
          stopAiImgRemote();
          return null;
        }
      }
    } catch (ConnectException e) {
      WtMessageHandler.printToLogFile("Could not connect to server at: " + url);
      WtMessageHandler.printException(e);
      stopAiImgRemote();
    } catch (Exception e) {
      WtMessageHandler.showError(e);
      stopAiImgRemote();
    } finally {
      conn.disconnect();
    }
    return null;
  }
  
  private String runTtsInstruction_intern(String text, String filename) throws Throwable {
    if (text == null || text.trim().isEmpty()) {
      return null;
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("AiRemote: runTtsInstruction: Ask AI started! URL: " + url);
    }
    String urlParameters = "{"
/*        
        + "\"do_sample\": true, "
        + "\"duration_seconds\": " + 0 + ", "
        + "\"model_id\": \"" + ttsModel  + "\", "
        + "\"prompt_influence\": "+ 0 + ", "
        + "\"text\": \"" + text + "\"}";
*/
        + "\"input\": \"" + text + "\", "
        + "\"model\": \"" + ttsModel + "\"}";
//        + "\"voice\": \"" + ttsModel + "\"}";
    
    byte[] postData = urlParameters.getBytes(StandardCharsets.UTF_8);
    
    URL checkUrl;
    try {
/*      
      String sUrl = ttsUrl + (ttsUrl.endsWith("/") ? "" : "/") + ttsModel + "/";
      checkUrl = new URL(sUrl);
*/      
      checkUrl = new URL(ttsUrl);
    } catch (MalformedURLException e) {
      WtMessageHandler.showError(e);
      stopAiTtsRemote();
      return null;
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("AiRemote: runTtsInstruction: postData: " + urlParameters);
    }
    HttpURLConnection conn = getConnection(postData, checkUrl, ttsApiKey);
    try {
      if (conn.getResponseCode() == HttpURLConnection.HTTP_OK) {
        try (InputStream inputStream = conn.getInputStream()) {
          storeByteStream(inputStream, filename);
//          String out = readStream(inputStream, "utf-8");
//          WtMessageHandler.printToLogFile("TTS-Out: " + out);
          WtMessageHandler.printToLogFile("TTS-Out: " + filename);
          return filename;
        }
      } else {
        try (InputStream inputStream = conn.getErrorStream()) {
          String error = readStream(inputStream, "utf-8");
          WtMessageHandler.printToLogFile("Got error: " + error + " - HTTP response code " + conn.getResponseCode());
          WtMessageHandler.printToLogFile("Url: " + checkUrl.toString() + "\nurlParameters: " + urlParameters);
          stopAiTtsRemote();
          return null;
        }
      }
    } catch (ConnectException e) {
      WtMessageHandler.printToLogFile("Could not connect to server at: " + url);
      WtMessageHandler.printException(e);
      stopAiTtsRemote();
    } catch (Exception e) {
      WtMessageHandler.showError(e);
      stopAiTtsRemote();
    } finally {
      conn.disconnect();
    }
    return null;
  }
  
  private HttpURLConnection getConnection(byte[] postData, URL url, String apiKey) throws Throwable {
    int trials = 0;
    while (trials < REMOTE_TRIALS) {
      trials++;
      try {
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setDoOutput(true);
        conn.setInstanceFollowRedirects(false);
        conn.setRequestMethod("POST");
        conn.setRequestProperty("Content-Type", "application/json");
        conn.setRequestProperty("charset", "utf-8");
        conn.setRequestProperty("Authorization", apiKey);
        conn.setRequestProperty("Content-Length", Integer.toString(postData.length));
        
        try (DataOutputStream wr = new DataOutputStream(conn.getOutputStream())) {
          wr.write(postData);
        }
        return conn;
      } catch (Exception e) {
        if (trials >= REMOTE_TRIALS) {
          throw new RuntimeException(e);
        }
      }
    }
    return null;
  }

  private String removeSurroundingBrackets(String out, String org) throws Throwable {
    if (out.startsWith("{") && out.endsWith("}")) {
      if (!org.startsWith("{") || !org.endsWith("}")) {
        return out.substring(1, out.length() - 1);
      }
    }
    return out;
  }
  
  private String filterOutput (String out, String org, String instruction, boolean onlyOneParagraph) throws Throwable {
    out = removeSurroundingBrackets(out, org);
    out = out.replace("\n", "\r").replace("\r\r", "\r").replace("\\\"", "\"").trim();
    if (onlyOneParagraph) {
      String[] inst = instruction.split("[-.:!?]");
      String[] parts = out.split("\r");
      String firstPart = parts[0].trim();
      if (parts.length > 1 && (firstPart.endsWith(":") || firstPart.startsWith(inst[0].trim()))) {
        out = parts[1].trim();
      } else {
        out = firstPart;
      }
      out = removeSurroundingBrackets(out, org);
      if (out.contains(":") && (!org.contains(":") || out.trim().startsWith(inst[0].trim()))) {
        parts = out.split(":");
        if (parts.length > 1) {
          int n = 1;
          if (parts.length > 2 && out.trim().startsWith(instruction.trim())) {
            n = 2;
          }
          out = parts[n];
          for (int i = n + 1; i < parts.length; i++) {
            out += ":" + parts[i];
          }
          out = removeSurroundingBrackets(out.trim(), org);
        }
      }
    }
    if (out.contains("->") && !org.contains("->") && out.startsWith(org)) {
//      WtMessageHandler.printToLogFile("Out(in): " + out);
      String vOut[] = out.split("->");
      if (vOut.length > 1) {
        out = vOut[1].trim();
      }
//      WtMessageHandler.printToLogFile("Out(out): " + out);
    }
    return out;
  }
  
  private String readStream(InputStream stream, String encoding) throws Throwable {
    StringBuilder sb = new StringBuilder();
    try (InputStreamReader isr = new InputStreamReader(stream, encoding);
         BufferedReader br = new BufferedReader(isr)) {
      String line;
      while ((line = br.readLine()) != null) {
        sb.append(line).append('\r');
      }
    }
    return sb.toString();
  }
  
  private void storeByteStream(InputStream inp, String filename) throws Throwable {
    try (OutputStream outp = new FileOutputStream(filename)) {
      byte[] bytes = new byte[BUFFER_SIZE];
      int rBytes = BUFFER_SIZE;
      while (rBytes == BUFFER_SIZE) {
        rBytes = inp.read(bytes);
        outp.write(bytes, 0, rBytes);
      }
      outp.flush();
      outp.close();
    }
  }
  
  private String parseJasonOutput(String text) throws Throwable {
    try {
      JSONObject jsonObject = new JSONObject(text);
      JSONArray choices;
      try {
        choices = jsonObject.getJSONArray("choices");
      } catch (Throwable t) {
        String error = jsonObject.getString("error");
        WtMessageHandler.showMessage(error);
        return null;
      }
      String content;
      JSONObject choice = choices.getJSONObject(0);
      if (aiType == AiType.CHAT) {
        JSONObject message = choice.getJSONObject("message");
        content = message.getString("content");
      } else {
        content = choice.getString("text");
      }
      if (debugMode) {
        WtMessageHandler.printToLogFile("AiRemote: parseJasonOutput: text: " + text);
        WtMessageHandler.printToLogFile("AiRemote: parseJasonOutput: content: " + content);
      }
      try {
        JSONObject contentObject = new JSONObject(content);
        try {
          int nLastObj = contentObject.length() - 1;
          Set<String> keySet = contentObject.keySet();
          int i = 0;
          for (String key : keySet) {
            if ( i == nLastObj) {
              content = getString(new JSONObject(contentObject.getString(key)));
              break;
            }
            i++;
          }
        } catch (Throwable t) {
          try {
            content = getString(contentObject);
          } catch (Throwable t1) {
            return content;
          }
        }
      } catch (Throwable t) {
        return content;
      }
      return content;
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      WtMessageHandler.showMessage(text);
      return null;
    }
  }
  
  private String parseJasonImgOutput(String text) throws Throwable {
    try {
      JSONObject jsonObject = new JSONObject(text);
      JSONArray data;
      try {
        data = jsonObject.getJSONArray("data");
      } catch (Throwable t) {
        String error = jsonObject.getString("error");
        WtMessageHandler.showMessage(error);
        return null;
      }
      JSONObject choice = data.getJSONObject(0);
      String url = choice.getString("url");
      if (debugMode) {
        WtMessageHandler.printToLogFile("AiRemote: parseJasonOutput: text: " + text);
        WtMessageHandler.printToLogFile("AiRemote: parseJasonOutput: url: " + url);
      }
      return url;
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      WtMessageHandler.showMessage(text);
      return null;
    }
  }
  
  private String getString(JSONObject contentObject) throws Throwable {
    try {
      JSONObject subObject = new JSONObject(contentObject);
      return subObject.toString();
    } catch (Throwable t) {
    }
    return contentObject.toString();
  }
  
  private void stopAiRemote() throws Throwable {
    config.setUseAiSupport(false);
    if (documents.getAiCheckQueue() != null) {
      documents.getAiCheckQueue().setStop();
      documents.setAiCheckQueue(null);
    }
    WtMessageHandler.showMessage(messages.getString("loAiServerConnectionError"));
  }
  
  private void stopAiImgRemote() throws Throwable {
    config.setUseAiImgSupport(false);
    WtMessageHandler.showMessage(messages.getString("loAiServerConnectionError"));
  }
  
  private void stopAiTtsRemote() throws Throwable {
    config.setUseAiTtsSupport(false);
    WtMessageHandler.showMessage(messages.getString("loAiServerConnectionError"));
  }
  
  public static String getInstruction(String mess, Locale locale) throws Throwable {
    if (locale == null || locale.Language == null || locale.Language.isEmpty()) {
      locale = new Locale("en", "US", "");
    }
//    ResourceBundle messages = WtOfficeTools.getMessageBundle(WtDocumentsHandler.getLanguage(new Locale("en", "", "")));
//    String instruction = messages.getString(mess) + " (language: " + locale.Language + ")"; 
    String instruction = mess + " (language: " + locale.Language + ")"; 
    return instruction;
  }
  
  public static String getLanguageName(Locale locale) throws Throwable {
    String lang = (locale == null || locale.Language == null || locale.Language.isEmpty()) ? "en" : locale.Language;
    Locale langLocale = new Locale(lang, "", "");
    return WtDocumentsHandler.getLanguage(langLocale).getName();
  }
  
  public static String addLanguageName(String instruction, Locale locale) throws Throwable {
    String langName = getLanguageName(locale);
    return instruction + " (language - " + langName + ")";
  }
  
}
