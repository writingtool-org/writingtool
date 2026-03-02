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
package org.writingtool.tools;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.ConnectException;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

/**
 * class to get information about and handle updates of WritingTool
 * @since 26.4
 * @author Fred Kruse
 */
public class WtUpdate {
  
  private final static String WT_RELEASE_URL = "https://writingtool.org/writingtool/releases/";
  private final static String WT_FILENAME_PREFIX = "WritingTool";
  private final static String WT_FILENAME_POSTFIX = "oxt";
  private WTReleaseVersion lastReleaseVersion;
  
  public WtUpdate() {
    try {
      lastReleaseVersion = getLastRelease();
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
      lastReleaseVersion = new WTReleaseVersion();
    }
  }
  
  public String getLastVersionFilename() {
    return lastReleaseVersion.fileName.substring(0, lastReleaseVersion.fileName.length() - WT_FILENAME_POSTFIX.length() - 1);
  }
  
  public String getLastVersionFullFilename() {
    return lastReleaseVersion.fileName;
  }
  
  public String getLastVersionUrl() {
    return WT_RELEASE_URL + lastReleaseVersion.fileName;
  }
  
  public String getReleaseDirUrl() {
    return WT_RELEASE_URL;
  }
  
  public String getLastVersionIdentifier() {
    return lastReleaseVersion.year + "." + lastReleaseVersion.month 
        + (lastReleaseVersion.version > 0 ? "." + lastReleaseVersion.version : "");
  }
  
  public boolean isNewVersion(String version) {
    WtMessageHandler.printToLogFile("WtUpdate: isNewVersion: last Release Version: " + lastReleaseVersion.fileName
        + ", Running Version: " + version);
    String[] vVersion = version.split("-");
    vVersion = vVersion[0].split("\\.");
    if (vVersion.length < 2 || vVersion.length > 3) {
      return false;
    }
    int year = Integer.parseInt(vVersion[0]);
    int month = Integer.parseInt(vVersion[1]);
    int iVersion = vVersion.length == 3 ? Integer.parseInt(vVersion[2]) : 0;
    if (lastReleaseVersion.year > year) {
      return true;
    } else if (lastReleaseVersion.year == year) {
      if (lastReleaseVersion.month > month) {
        return true;
      } else if (lastReleaseVersion.month == month) {
        if (lastReleaseVersion.version > iVersion) {
          return true;
        }
      }
      
    }
    return false;
  }
  
  private HttpURLConnection getConnection(URL url) throws Throwable {
    try {
      HttpURLConnection conn = (HttpURLConnection) url.openConnection();
      conn.setDoOutput(true);
      conn.setInstanceFollowRedirects(true);
      conn.setRequestMethod("GET");
      conn.setConnectTimeout(0);
      conn.setReadTimeout(0);
      return conn;
    } catch (Exception e) {
      throw new RuntimeException(e);
    }
  }

  private List<String> readStreamToList(InputStream stream, String encoding) throws Throwable {
    List<String> out = new ArrayList<>(); 
    if (stream != null) {
      try (InputStreamReader isr = new InputStreamReader(stream, encoding);
           BufferedReader br = new BufferedReader(isr)) {
        String line;
        while ((line = br.readLine()) != null) {
          if (line.contains(WT_FILENAME_PREFIX) && line.contains(WT_FILENAME_POSTFIX)) {
            line = line.replaceAll(".*" + WT_FILENAME_PREFIX, WT_FILENAME_PREFIX).replaceAll(WT_FILENAME_POSTFIX + ".*", WT_FILENAME_POSTFIX);
//            WtMessageHandler.printToLogFile("Read Dir: " + line.toString());
            out.add(line.toString());
          }
        }
      } catch (IOException e) {
        //  stream is suddenly closed -> return null
        WtMessageHandler.printToLogFile("WtUpdate: readStreamToList: IOException: " + e.getMessage());
        return null;
      }
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
        WtMessageHandler.printToLogFile("WtUpdate: readStream: IOException: " + e.getMessage());
        return null;
      }
    }
    return sb.toString();
  }
  
  private WTReleaseVersion fileNameToReleaseVersion(String filename) {
    WTReleaseVersion releaseVersion = new WTReleaseVersion();
    if(!filename.endsWith("." + WT_FILENAME_POSTFIX)) {
      return null;
    }
    String fname = filename.substring(0, filename.length() - WT_FILENAME_POSTFIX.length() - 1);
    String[] name = fname.split("-");
    if(name.length != 2) {
      return null;
    }
    String[] version = name[1].split("\\.");
    if (version.length < 2 || version.length > 3) {
      return null;
    }
    releaseVersion.fileName = filename;
    releaseVersion.year = Integer.parseInt(version[0]);
    releaseVersion.month = Integer.parseInt(version[1]);
    if (version.length == 3) {
      releaseVersion.version = Integer.parseInt(version[2]);
    } else {
      releaseVersion.version = 0;
    }
    return releaseVersion;
  }
  
  private WTReleaseVersion getLastReleaseVersion(List<String> fileNames) {
    WTReleaseVersion lastVersion = new WTReleaseVersion();
    for (String fileName : fileNames) {
      WTReleaseVersion rVersion = fileNameToReleaseVersion(fileName);
      if (rVersion.year > lastVersion.year || (rVersion.year == lastVersion.year && rVersion.month > lastVersion.month) 
          || (rVersion.year == lastVersion.year && rVersion.month == lastVersion.month && rVersion.version > lastVersion.version)) {
        lastVersion.year = rVersion.year;
        lastVersion.month = rVersion.month;
        lastVersion.version = rVersion.version;
        lastVersion.fileName = rVersion.fileName;
      }
    }
    return lastVersion;
  }
  
  private WTReleaseVersion getLastRelease() throws Throwable {
    
    URL dirUrl;
    try {
      dirUrl = new URL(WT_RELEASE_URL);
    } catch (MalformedURLException e) {
      WtMessageHandler.showError(e);
      return null;
    }
    HttpURLConnection conn = null;
    try {
      conn = getConnection(dirUrl);
    } catch (RuntimeException e) {
      WtMessageHandler.showError(e);
      return null;
    }
    try {
      if (conn.getResponseCode() == HttpURLConnection.HTTP_OK) {
        try (InputStream inputStream = conn.getInputStream()) {
          List<String> outList = readStreamToList(inputStream, "utf-8");
          if (outList == null) {
            return null;
          }
          return getLastReleaseVersion(outList);
        }
      } else {
        try (InputStream inputStream = conn.getErrorStream()) {
          int responseCode = conn.getResponseCode();
          String error = readStream(inputStream, "utf-8");
          String msg = "Got error: " + error + " - HTTP response code " + responseCode;
          if (responseCode == 404) {
              WtMessageHandler.showMessage(msg);
          } else {
            WtMessageHandler.showMessage(msg);
          }
          return null;
        }
      }
    } catch (ConnectException e) {
      WtMessageHandler.showError(e);
    } catch (Exception e) {
      WtMessageHandler.showError(e);
    } finally {
      conn.disconnect();
    }
    return null;
  }
  
  private class WTReleaseVersion {
    public int year = 0;
    public int month = 0;
    public int version = 0;
    public String fileName = null;
  }


}
