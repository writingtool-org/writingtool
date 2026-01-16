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

import java.io.BufferedReader;
import java.io.DataOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.ConnectException;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.ResourceBundle;
import java.util.Set;

import org.json.JSONArray;
import org.json.JSONObject;
import org.writingtool.WtDocumentsHandler;
import org.writingtool.config.WtConfiguration;
import org.writingtool.dialogs.WtAiDialog;
import org.writingtool.sidebar.WtSidebarContent;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;

import com.sun.star.lang.Locale;


/**
 * Class to communicate with a AI API
 * @since 1.0
 * @author Fred Kruse
 */
public class WtAiRemote {
  
  private final static int REMOTE_TRIALS = 5;
  private final static int BUFFER_SIZE = 20000;
  private final static int CONNECT_TIMEOUT = 0;
  private final static int READ_TIMEOUT = 0;
  private final static int WAIT_TIMEOUT = 1000; //  Hundredth of a second
  
  private static final ResourceBundle messages = WtOfficeTools.getMessageBundle();

  public final static int MAX_TEXT_LENGTH = 20000;
  public final static int MAX_OID = 32000;

  public final static String CORRECT_INSTRUCTION = "Output the grammatically and orthographically corrected text without comments";
  public final static String STYLE_INSTRUCTION = "Output the stylistic reformulated text without comments";
  public final static String REFORMULATE_INSTRUCTION = "Output the rephrased text without comments";
  public final static String EXPAND_INSTRUCTION = "Output the expanded text";
  public final static String SYNONYMS_INSTRUCTION = "List at least 3 and a maximum of 20 synonyms of the following word, without comments";
  
  public final static float CORRECT_TEMPERATURE = 0.0f;
  public final static float REFORMULATE_TEMPERATURE = 0.4f;
  public final static float EXPAND_TEMPERATURE = 0.7f;
  public final static float SYNONYM_TEMPERATURE = 0.0f;
  
  public static enum AiCategory { Text, Image, Speech };
  public static enum AiCommand { CorrectGrammar, ImproveStyle, ReformulateText, ExpandText, SynonymsOfWord, GeneralAi };
  
  private static boolean isRunning = false;
  private static Map<Integer, String> results = new HashMap<>();
  private static List<AiEntry> entries = new ArrayList<>();
  private static int lastOid = 0;
  private static boolean hasPrintedInfo = false;

  private enum AiType { EDITS, COMPLETIONS, CHAT, GENERATE }
  
  private boolean debugModeTm = WtOfficeTools.DEBUG_MODE_TM;
  private int debugMode = WtOfficeTools.DEBUG_MODE_AI;

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
  
  private String outText;
  private boolean isDone;
  
  private int oId = 0;
  
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
    } else if (url.endsWith("/generate/") || url.endsWith("/generate")) {
      aiType = AiType.GENERATE;
    } else {
      aiType = AiType.CHAT;
    }
  }
  
  private int getId() {
    if (lastOid == MAX_OID) {
      lastOid = 0;
    }
    lastOid++;
    return lastOid;
  }

  public String runInstruction(String instruction, String text, float temperature, 
      int seed, Locale locale, boolean onlyOneParagraph, boolean preferred) throws Throwable {
    return runInstructionGeneral(AiCategory.Text, instruction, text, null, null, temperature, seed, 0, 0, 0, 
        locale, onlyOneParagraph, preferred);
  }

  public String runImgInstruction(String instruction, String exclude, int step, int seed, int height, 
      int width, boolean preferred) throws Throwable {
    return runInstructionGeneral(AiCategory.Image, instruction, null, exclude, null, 0, seed, step, height, width, 
        null, false, preferred);
  }

  public String runTtsInstruction(String text, String filename, boolean preferred) throws Throwable {
    return runInstructionGeneral(AiCategory.Speech, null, text, null, filename, 0, 0, 0, 0, 0, null, false, preferred);
  }
  
  public String runInstructionGeneral(AiCategory category, String instruction, String text, String exclude, String filename, 
      float temperature, int seed, int step, int height, int width, 
      Locale locale, boolean onlyOneParagraph, boolean preferred) throws Throwable {
    if (oId > 0) {
      throw new RuntimeException("Duplicate OID in WtAiRemote");
    }
    oId = getId();
    AiEntry entry = new AiEntry(category, oId, instruction, text, exclude, filename, temperature, seed, 
        step, height, width, locale, onlyOneParagraph);
    if (preferred) {
      entries.add(0, entry);
    } else {
      entries.add(entry);
    }
    wakeRunAiEntries();
    try {
      while (!results.containsKey(oId)) {
        try {
          Thread.sleep(100);
        } catch (InterruptedException e) {
          WtMessageHandler.printException(e);
        }
      }
      String result = results.get(oId);
      results.remove(oId);
      return result;
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      return null;
    }
  }
/*  
  public String runImgInstruction(String instruction, String exclude, int step, int seed, int size, boolean preferred) throws Throwable {
    oId = getId();
    AiEntry entry = new AiEntry(oId, instruction, exclude, step, seed, size);
    if (preferred) {
      entries.add(0, entry);
    } else {
      entries.add(entry);
    }
    wakeRunAiEntries();
    try {
      while (!results.containsKey(oId)) {
        try {
          Thread.sleep(100);
        } catch (InterruptedException e) {
          WtMessageHandler.printException(e);
        }
      }
      String result = results.get(oId);
      results.remove(oId);
      return result;
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      return null;
    }
  }

  public String runTtsInstruction(String text, String filename, boolean preferred) throws Throwable {
    oId = getId();
    AiEntry entry = new AiEntry(oId, text, filename);
    if (preferred) {
      entries.add(0, entry);
    } else {
      entries.add(entry);
    }
    wakeRunAiEntries();
    try {
      while (!results.containsKey(oId)) {
        try {
          Thread.sleep(100);
        } catch (InterruptedException e) {
          WtMessageHandler.printException(e);
        }
      }
      String result = results.get(oId);
      results.remove(oId);
      return result;
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      return null;
    }
  }
*/  
  private String runInstruction_intern(String instruction, String orgText, float temperature, 
      int seed, Locale locale, boolean onlyOneParagraph) throws Throwable {
    if (instruction == null || orgText == null) {
      return null;
    }
    instruction = instruction.trim();
    String text = orgText.trim();
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
    if (debugMode > 1) {
      WtMessageHandler.printToLogFile("AiRemote: runInstruction: Ask AI started! URL: " + url);
    }
//    String langName = getLanguageName(locale);
    String langName = locale.Language;
    String org = text == null ? instruction : text;
    if (text != null) {
      text = text.replace("\n", " ").replace("\r", " ").replace("\"", "\\\"").replace("\t", " ");
    }
    instruction = instruction.replace("\n", " ").replace("\r", " ").replace("\"", "\\\"").replace("\t", " ");
    String urlParameters;
//    instruction = addLanguageName(instruction, locale);
    if (aiType == AiType.CHAT) {
      urlParameters = "{\"model\": \"" + model + "\", " 
//        + "\"response_format\": { \"type\": \"json_object\" }, "
          + "\"stream\": false, "
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
          + "\"prompt\": \"" + instruction + ": {" + text + "}\", "
          + (aiType == AiType.GENERATE ? "\"stream\": false, " : "")
          + (seed > 0 ? "\"seed\": " + seed + ", " : "")
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
    if (debugMode > 1) {
      WtMessageHandler.printToLogFile("AiRemote: runInstruction: postData: " + urlParameters);
    }
    HttpURLConnection conn = null;
    int trials = 0;
    while (trials < REMOTE_TRIALS) {
      trials++;
      try {
        conn = getConnection(postData, checkUrl, apiKey);
      } catch (RuntimeException e) {
        WtMessageHandler.printException(e);
        stopAiRemote();
        return null;
      }
      try {
        if (conn.getResponseCode() == HttpURLConnection.HTTP_OK) {
          try (InputStream inputStream = conn.getInputStream()) {
            List<String> outList = readStreamToList(inputStream, "utf-8");
            if (outList == null) {
              return null;
            }
            String out = parseJasonOutputOllama(outList);
            if (out == null) {
              out = parseJasonOutput(outList);
            }
            if (out == null) {
              return null;
            }
            out = filterOutput (out, org, instruction, onlyOneParagraph);
            if (debugModeTm) {
              long runTime = System.currentTimeMillis() - startTime;
              WtMessageHandler.printToLogFile("AiRemote: runInstruction: Time to generate Answer: " + runTime);
            }
            WtSidebarContent sidebarContent = documents.getSidebarContent();
            if (sidebarContent != null) {
              sidebarContent.setTextToAiResultBox(orgText, out, instruction);
            }
            return out;
          }
        } else {
          try (InputStream inputStream = conn.getErrorStream()) {
            int responseCode = conn.getResponseCode();
            String error = readStream(inputStream, "utf-8");
            String msg = "Got error: " + error + " - HTTP response code " + responseCode
                + "\nurlParameters: " + urlParameters;
            if (responseCode == 404) {
              WtMessageHandler.printToLogFile("Could not connect to server at: " + url);
              WtMessageHandler.printToLogFile(msg);
              stopAiRemote();
            } else {
              WtMessageHandler.showMessage(msg);
            }
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
  
  private String runImgInstruction_intern(String instruction, String exclude, int step, int seed, int height, int width) throws Throwable {
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
/*    
    if (size != 128 && size != 256 && size != 512) {
      size = 256;
    }
*/
    if (debugMode > 1) {
      WtMessageHandler.printToLogFile("AiRemote: runImgInstruction: Ask AI started! URL: " + url);
    }
    String urlParameters = "{\"model\": \"" + imgModel + "\", " 
        + "\"prompt\": \"" + instruction + "\", "
        + (seed > 0 ? "\"seed\": " + seed + ", " : "")
        + "\"size\": \"" + width + "x" + height + "\", "
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
    if (debugMode > 1) {
      WtMessageHandler.printToLogFile("AiRemote: runImgInstruction: postData: " + urlParameters);
    }
    HttpURLConnection conn;
    try {
      conn = getConnection(postData, checkUrl, imgApiKey);
    } catch (RuntimeException e) {
      WtMessageHandler.printException(e);
      stopAiImgRemote();
      return null;
    }
    try {
      if (conn.getResponseCode() == HttpURLConnection.HTTP_OK) {
        try (InputStream inputStream = conn.getInputStream()) {
          String out = readStream(inputStream, "utf-8");
          if (out == null) {
            return null;
          }
          return parseJasonImgOutput(out);
        }
      } else {
        try (InputStream inputStream = conn.getErrorStream()) {
          int responseCode = conn.getResponseCode();
          String error = readStream(inputStream, "utf-8");
          String msg = "Got error: " + error + " - HTTP response code " + responseCode
              + "\nurlParameters: " + urlParameters;
          if (responseCode == 404) {
            WtMessageHandler.printToLogFile("Could not connect to server at: " + url);
            WtMessageHandler.printToLogFile(msg);
            stopAiImgRemote();
          } else {
            WtMessageHandler.showMessage(msg);
          }
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
    text = text.replace("\n", " ").replace("\r", " ").replace("\"", "\\\"").replace("\t", " ");
    if (debugMode > 1) {
      WtMessageHandler.printToLogFile("AiRemote: runTtsInstruction: Ask AI started! URL: " + url);
    }
    String urlParameters = "{"
        + "\"input\": \"" + text + "\", "
        + "\"model\": \"" + ttsModel + "\"}";
    
    byte[] postData = urlParameters.getBytes(StandardCharsets.UTF_8);
    
    URL checkUrl;
    try {
      checkUrl = new URL(ttsUrl);
    } catch (MalformedURLException e) {
      WtMessageHandler.showError(e);
      stopAiTtsRemote();
      return null;
    }
    if (debugMode > 1) {
      WtMessageHandler.printToLogFile("AiRemote: runTtsInstruction: postData: " + urlParameters);
    }
    HttpURLConnection conn;
    try {
      conn = getConnection(postData, checkUrl, ttsApiKey);
    } catch (RuntimeException e) {
      WtMessageHandler.printException(e);
      stopAiTtsRemote();
      return null;
    }
    try {
      if (conn.getResponseCode() == HttpURLConnection.HTTP_OK) {
        try (InputStream inputStream = conn.getInputStream()) {
          storeByteStream(inputStream, filename);
          WtMessageHandler.printToLogFile("TTS-Out: " + filename);
          return filename;
        }
      } else {
        try (InputStream inputStream = conn.getErrorStream()) {
          int responseCode = conn.getResponseCode();
          String error = readStream(inputStream, "utf-8");
          String msg = "Got error: " + error + " - HTTP response code " + responseCode
              + "\nurlParameters: " + urlParameters;
          if (responseCode == 404) {
            WtMessageHandler.printToLogFile("Could not connect to server at: " + url);
            WtMessageHandler.printToLogFile(msg);
            stopAiRemote();
          } else {
            WtMessageHandler.showMessage(msg);
          }
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
        conn.setInstanceFollowRedirects(true);
        conn.setRequestMethod("POST");
        conn.setRequestProperty("Content-Type", "application/json");
        conn.setRequestProperty("charset", "utf-8");
        conn.setRequestProperty("Authorization", apiKey);
        conn.setRequestProperty("Content-Length", Integer.toString(postData.length));
        int connectTimeout = conn.getConnectTimeout();
        int readTimeout = conn.getReadTimeout();
        if (!hasPrintedInfo) {
          WtMessageHandler.printToLogFile("AI-Host: " + url.getHost());
          WtMessageHandler.printToLogFile("connectTimeout: " + connectTimeout);
          WtMessageHandler.printToLogFile("readTimeout: " + readTimeout);
          hasPrintedInfo = true;
        }
        conn.setConnectTimeout(CONNECT_TIMEOUT);
        conn.setReadTimeout(READ_TIMEOUT);
        
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
  
  public static String removeSurroundingQuotes(String out, String org) throws Throwable {
    out = out.trim();
    org = org.trim();
    if (out.startsWith("\"") && out.endsWith("\"")) {
      if (!org.startsWith("\\\"") || !org.endsWith("\"")) {
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
    if (stream != null) {
      try (InputStreamReader isr = new InputStreamReader(stream, encoding);
           BufferedReader br = new BufferedReader(isr)) {
        String line;
        while ((line = br.readLine()) != null) {
          sb.append(line).append('\r');
        }
      } catch (IOException e) {
        //  stream is suddenly closed -> return null
        if (debugMode > 0) {
          WtMessageHandler.printToLogFile("AiRemote: readStream: IOException: " + e.getMessage());
        }
        return null;
      }
    }
    return sb.toString();
  }
  
  private String listToString(List<String> list) throws Throwable {
    if (list == null) {
      return null;
    }
    StringBuilder sb = new StringBuilder();
    for(int i = 0; i < list.size(); i++) {
      sb.append(list.get(i)).append('\r');
    }
    return sb.toString();
  }
  
  private List <String> readStreamToList(InputStream stream, String encoding) throws Throwable {
    List<String> out = new ArrayList<>(); 
    if (stream != null) {
      try (InputStreamReader isr = new InputStreamReader(stream, encoding);
           BufferedReader br = new BufferedReader(isr)) {
        String line;
        while ((line = br.readLine()) != null) {
          out.add(line.toString());
        }
      } catch (IOException e) {
        //  stream is suddenly closed -> return null
        if (debugMode > 0) {
          WtMessageHandler.printToLogFile("AiRemote: readStream: IOException: " + e.getMessage());
        }
        return null;
      }
    }
    return out;
  }
  
  private void storeByteStream(InputStream inp, String filename) throws Throwable {
    try (OutputStream outp = new FileOutputStream(filename)) {
      byte[] bytes = new byte[BUFFER_SIZE];
      int rBytes;
      while ((rBytes = inp.read(bytes)) > -1) {
        outp.write(bytes, 0, rBytes);
      }
      outp.flush();
      outp.close();
      inp.close();
    }
  }
  
  private String parseJasonOutput(String text) throws Throwable {
    try {
      JSONObject jsonObject = new JSONObject(text);
      JSONArray choices;
      if (debugMode > 2) {
        WtMessageHandler.printToLogFile("AiRemote: parseJasonOutput: text: " + text);
      }
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
      if (debugMode > 1) {
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
  
  private String parseJasonOutput(List<String> textList) throws Throwable {
    return parseJasonOutput(listToString(textList));
  }
  
  private String parseJasonOutputOllama(List<String> textList) throws Throwable {
    if (debugMode > 2) {
      WtMessageHandler.printToLogFile("AiRemote: parseJasonOutput: text: " + textList.get(0));
    }
    if (textList == null) {
      return null;
    }
    if (textList.isEmpty()) {
      return "";
    }
    StringBuilder sb = new StringBuilder();
    boolean done = false;
    boolean hasMessage = true;
    for(int i = 0; i < textList.size() && !done; i++) {
      JSONObject jsonObject = new JSONObject(textList.get(i));
      if (i == 0) {
        if (!jsonObject.has("done")) {
          return null;
        }
        if (!jsonObject.has("message")) {
          if (!jsonObject.has("response")) {
            return null;
          }
          hasMessage = false;
        }
      }
      String content;
      if (hasMessage) {
        JSONObject message = jsonObject.getJSONObject("message");
        content = message.getString("content");
      } else {
        content = jsonObject.getString("response");
      }
      if (debugMode > 2) {
        WtMessageHandler.printToLogFile("AiRemote: parseJasonOutput: text: " + textList.get(i));
      }
      done = jsonObject.getBoolean("done");
      sb.append(content);
    }
    return sb.toString();
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
      if (debugMode > 1) {
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
    documents.getSidebarContent().setAiSupport(config.useAiSupport());
    if (documents.getAiCheckQueue() != null) {
      documents.getAiCheckQueue().setStop();
      documents.setAiCheckQueue(null);
    }
    WtMessageHandler.showMessage(messages.getString("loAiServerConnectionError"));
    WtAiDialog aiDialog = WtAiParagraphChanging.getAiDialog();
    if (aiDialog != null) {
      aiDialog.closeDialog();
    } 
  }
  
  private void stopAiImgRemote() throws Throwable {
    config.setUseAiImgSupport(false);
    documents.getSidebarContent().setAiSupport(config.useAiSupport());
    WtMessageHandler.showMessage(messages.getString("loAiServerConnectionError"));
    WtAiDialog aiDialog = WtAiParagraphChanging.getAiDialog();
    if (aiDialog != null) {
      aiDialog.closeDialog();
    } 
  }
  
  private void stopAiTtsRemote() throws Throwable {
    config.setUseAiTtsSupport(false);
    documents.getSidebarContent().setAiSupport(config.useAiSupport());
    WtMessageHandler.showMessage(messages.getString("loAiServerConnectionError"));
    WtAiDialog aiDialog = WtAiParagraphChanging.getAiDialog();
    if (aiDialog != null) {
      aiDialog.closeDialog();
    } 
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
  
  private String runInstructionText(String instruction, String orgText, float temperature, 
      int seed, Locale locale, boolean onlyOneParagraph) {
    outText = null;
    isDone = false;
    Thread t = new Thread(new Runnable() {
      public void run() {
        try {
          outText = runInstruction_intern(instruction, orgText, temperature, seed, locale, onlyOneParagraph);
          isDone = true;
        } catch (Throwable e) {
          WtMessageHandler.showError(e);
        }
      }
    });
    t.start();
    int n = 0;
    while (!isDone && n < WAIT_TIMEOUT) {
      try {
        Thread.sleep(10);
      } catch (InterruptedException e) {
        WtMessageHandler.printException(e);
      }
    }
    return outText;
  }
  
  private String runInstructionImage(String instruction, String exclude, int step, int seed, int height, int width) {
    outText = null;
    isDone = false;
    Thread t = new Thread(new Runnable() {
      public void run() {
        try {
          outText = runImgInstruction_intern(instruction, exclude, step, seed, height, width);
          isDone = true;
        } catch (Throwable e) {
          WtMessageHandler.showError(e);
        }
      }
    });
    t.start();
    int n = 0;
    while (!isDone && n < WAIT_TIMEOUT) {
      try {
        Thread.sleep(10);
      } catch (InterruptedException e) {
        WtMessageHandler.printException(e);
      }
    }
    return outText;
  }
  
  private String runInstructionTTS(String text, String filename) {
    outText = null;
    isDone = false;
    Thread t = new Thread(new Runnable() {
      public void run() {
        try {
          outText = runTtsInstruction_intern(text, filename);
          isDone = true;
        } catch (Throwable e) {
          WtMessageHandler.showError(e);
        }
      }
    });
    t.start();
    int n = 0;
    while (!isDone && n < WAIT_TIMEOUT) {
      try {
        Thread.sleep(10);
      } catch (InterruptedException e) {
        WtMessageHandler.printException(e);
      }
    }
    return outText;
  }
  
  private void wakeRunAiEntries() {
    if (!isRunning) {
      runAiEntries();
    }
  }
  
  private void runAiEntries() {
    Thread t = new Thread(new Runnable() {
      public void run() {
        try {
          isRunning = true;
          while (entries.size() > 0) {
            AiEntry entry = entries.get(0);
            entries.remove(0);
            String result = null;
            if (entry.category == AiCategory.Text && config.useAiSupport()) {
              result = runInstructionText(entry.instruction, entry.text, entry.temperature, entry.seed, entry.locale, entry.onlyOneParagraph);
            } else if (entry.category == AiCategory.Image && config.useAiImgSupport()) {
              result = runInstructionImage(entry.instruction, entry.exclude, entry.step, entry.seed, entry.height, entry.width);
            } else if (entry.category == AiCategory.Speech && config.useAiTtsSupport()) {
              result = runInstructionTTS(entry.text, entry.filename);
            }
            results.put(entry.oId, result);
          }
          isRunning = false;
        } catch (Throwable e) {
          WtMessageHandler.showError(e);
        }
      }
    });
    t.start();
  }

  private class AiEntry {
    public final AiCategory category;
    public final int oId;
    public final String instruction;
    public final String text;
    public final String exclude;
    public final String filename;
    public final float temperature;
    public final int seed;
    public final int step;
    public final int height;
    public final int width;
    public final Locale locale;
    public final boolean onlyOneParagraph;
    
    AiEntry (AiCategory category, int oId, String instruction, String text, String exclude, String filename, float temperature, 
        int seed, int step, int height, int width, Locale locale, boolean onlyOneParagraph) {
      this.category = category;
      this.oId = oId;
      this.instruction = instruction;
      this.text = text;
      this.exclude = exclude;
      this.filename = filename;
      this.temperature = temperature;
      this.seed = seed;
      this.step = step;
      this.height = height;
      this.width = width;
      this.locale = locale;
      this.onlyOneParagraph = onlyOneParagraph;
    }
/*    
    AiEntry(int oId, String instruction, String text, float temperature, int seed, Locale locale, boolean onlyOneParagraph) {
      this (AiCategory.Text, oId, instruction, text, null, null, temperature, seed, 0, 0, locale, onlyOneParagraph);
    }
    
    AiEntry(int oId, String instruction, String exclude, int step, int seed, int size) {
      this (AiCategory.Image, oId, instruction, null, exclude, null, 0, seed, step, size, null, false);
    }
    
    AiEntry(int oId, String text, String filename) {
      this (AiCategory.Speech, oId, null, text, null, filename, 0, 0, 0, 0, null, false);
    }
*/  
  }
  
}
