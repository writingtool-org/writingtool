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
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InvalidClassException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.io.Serializable;
import java.net.URI;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.zip.GZIPInputStream;
import java.util.zip.GZIPOutputStream;

import org.writingtool.WtDocumentCache.SerialLocale;
import org.writingtool.WtIgnoredMatches.LocaleEntry;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtOfficeTools.LoErrorType;
import org.writingtool.tools.WtVersionInfo;

import com.sun.star.frame.XModel;
import com.sun.star.lang.Locale;
import com.sun.star.lang.XComponent;
import com.sun.star.uno.UnoRuntime;

/**
 * Class to read and write LT-Office-Extension-Cache
 * @since 1.0
 * @author Fred Kruse
 */
public class WtCacheIO implements Serializable {

  private static final long serialVersionUID = 1L;
  private static final boolean DEBUG_MODE = WtOfficeTools.DEBUG_MODE_IO;

  private static final long MAX_CACHE_TIME = 365 * 24 * 3600000;      //  Save cache files maximal one year
  private static final String CACHEFILE_MAP = "WtCacheMap";           //  Name of cache map file
  private static final String CACHEFILE_PREFIX = "WtCache";           //  Prefix for cache files (simply a number is added for file name)
  private static final String CACHEFILE_EXTENSION = "lcz";            //  extension of the files name (Note: cache files are in zip format)
  private static final int MIN_CHARACTERS_TO_SAVE_CACHE = 25000;      //  Minimum characters of document for saving cache

  private static final String SPELL_CACHEFILE = "WtSpellCache." + CACHEFILE_EXTENSION;  //  Spell cache name
  
  private String documentPath = null;
  private AllCaches allCaches;
  
  WtCacheIO(XComponent xComponent) {
    setDocumentPath(xComponent);
  }
  
  WtCacheIO() {}
  
  void setDocumentPath(XComponent xComponent) {
    if (xComponent != null) {
      documentPath = getDocumentPath(xComponent);
    }
  }
  
  /** 
   * returns the text cursor (if any)
   * returns null if it fails
   */
  private static String getDocumentPath(XComponent xComponent) {
    try {
      XModel xModel = UnoRuntime.queryInterface(XModel.class, xComponent);
      if (xModel == null) {
        WtMessageHandler.printToLogFile("CacheIO: getDocumentPath: XModel not found!");
        return null;
      }
      String url = xModel.getURL();
      if (url == null || !url.startsWith("file://")) {
        if (url != null && !url.isEmpty()) {
          WtMessageHandler.printToLogFile("Not a file URL: " + (url == null ? "null" : url));
        }
        return null;
      }
      if (DEBUG_MODE) {
        WtMessageHandler.printToLogFile("CacheIO: getDocumentPath: file URL: " + url);
      }
      URI uri = new URI(url);
      return uri.getPath();
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown are printed
      return null;           // Return null as method failed
    }
  }

  /**
   * get the path to the cache file
   * if create == true: a new file is created if the file does not exist
   * if create == false: null is returned if the file does not exist
   */
  private String getCachePath(boolean create) {
    if (documentPath == null) {
      WtMessageHandler.printToLogFile("CacheIO: getCachePath: documentPath == null!");
      return null;
    }
    File cacheDir = WtOfficeTools.getCacheDir();
    if (cacheDir == null) {
      WtMessageHandler.printToLogFile("CacheIO: getCachePath: cacheDir == null!");
      return null;
    }
    if (DEBUG_MODE) {
      WtMessageHandler.printToLogFile("CacheIO: getCachePath: cacheDir: " + cacheDir.getAbsolutePath());
    }
    CacheFile cacheFile = new CacheFile(cacheDir);
    String cacheFileName = cacheFile.getCacheFileName(documentPath, create);
    if (cacheFileName == null) {
      WtMessageHandler.printToLogFile("CacheIO: getCachePath: cacheFileName == null!");
      return null;
    }
    File cacheFilePath = new File(cacheDir, cacheFileName);
    if (!create) {
      cacheFile.cleanUp(cacheFileName);
    }
    if (DEBUG_MODE) {
      WtMessageHandler.printToLogFile("CacheIO: getCachePath: cacheFilePath: " + cacheFilePath.getAbsolutePath());
    }
    return cacheFilePath.getAbsolutePath();
  }
  
  /**
   * save all caches (document cache, all result caches) to cache file
   */
  private void saveAllCaches(String cachePath) {
    try {
      GZIPOutputStream fileOut = new GZIPOutputStream(new FileOutputStream(cachePath));
      ObjectOutputStream out = new ObjectOutputStream(fileOut);
      out.writeObject(allCaches);
      out.close();
      fileOut.close();
      WtMessageHandler.printToLogFile("Caches saved to: " + cachePath);
      if (DEBUG_MODE) {
        printCacheInfo();
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown are printed
    }
  }
  
  /**
   * returns true if the number of characters of a document exceeds 
   * the minimal number of characters to save the cache
   */
  private boolean exceedsSaveSize(WtDocumentCache docCache) {
    int nChars = 0;
    for (int i = 0; i < docCache.size(); i++) {
      nChars += docCache.getFlatParagraph(i).length();
      if (nChars > MIN_CHARACTERS_TO_SAVE_CACHE) {
        return true;
      }
    }
    return false;
  }
  
  /**
   * save all caches if the document exceeds the defined minimum of paragraphs
   */
  public void saveCaches(WtDocumentCache docCache, List<WtResultCache> paragraphsCache, WtResultCache aiSuggestionCache,
      WtIgnoredMatches ignoredMatches, WtConfiguration config, WtDocumentsHandler mDocHandler) {
    String cachePath = getCachePath(true);
    if (cachePath != null) {
      try {
        if (!ignoredMatches.isEmpty() || exceedsSaveSize(docCache)) {
          allCaches = new AllCaches(docCache, paragraphsCache, aiSuggestionCache, 
              mDocHandler.getAllDisabledRules(), config.getDisabledRuleIds(), config.getDisabledCategoryNames(), 
              config.getEnabledRuleIds(), ignoredMatches, WtVersionInfo.ltVersion());
          saveAllCaches(cachePath);
        } else {
          File file = new File( cachePath );
          if (file.exists() && !file.isDirectory()) {
            file.delete();
          }
        }
      } catch (Throwable t) {
        WtMessageHandler.printToLogFile("CacheIO: saveCaches: " + t.getMessage());
        if (DEBUG_MODE) {
          WtMessageHandler.printException(t);     // all Exceptions thrown are printed
        }
      }
    }
  }
  
  /**
   * read all caches (document cache, all result caches) from cache file if it exists
   */
  public boolean readAllCaches(WtConfiguration config, WtDocumentsHandler mDocHandler) {
    String cachePath = getCachePath(false);
    if (cachePath == null) {
      return false;
    }
    try {
      File file = new File( cachePath );
      if (file.exists() && !file.isDirectory()) {
        GZIPInputStream fileIn = new GZIPInputStream(new FileInputStream(file));
        ObjectInputStream in = new ObjectInputStream(fileIn);
        allCaches = (AllCaches) in.readObject();
        in.close();
        fileIn.close();
        WtMessageHandler.printToLogFile("Caches read from: " + cachePath);
        if (DEBUG_MODE) {
          printCacheInfo();
        }
        if (runSameRules(config, mDocHandler)) {
          return true;
        } else {
          WtMessageHandler.printToLogFile("Version or active rules have changed: Cache rejected (Cache Version: " 
                + allCaches.ltVersion + ", actual LT Version: " + WtVersionInfo.ltVersion() + ")");
          return false;
        }
      }
    } catch (InvalidClassException e) {
      WtMessageHandler.printToLogFile("Old cache Version: Cache not read");
      return false;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown are printed
    }
    return false;
  }
  
  /**
   * Test if cache was created with same rules
   */
  private boolean runSameRules(WtConfiguration config, WtDocumentsHandler mDocHandler) {
    if (allCaches == null || allCaches.docCache == null || allCaches.docCache.toParaMapping.size() != WtDocumentCache.NUMBER_CURSOR_TYPES ) {
      if (DEBUG_MODE) {
        WtMessageHandler.printToLogFile("allCaches == null: " + (allCaches == null) + "; allCaches.docCache == null: " + (allCaches.docCache == null)
        + "; allCaches.docCache.toTextMapping.size(): " + (allCaches.docCache.toParaMapping.size()));
      }
      return false;
    }
    if (!allCaches.ltVersion.equals(WtVersionInfo.ltVersion())) {
      return false;
    }
    if (config.getEnabledRuleIds().size() != allCaches.enabledRuleIds.size() || config.getDisabledRuleIds().size() != allCaches.disabledRuleIds.size() 
          || config.getDisabledCategoryNames().size() != allCaches.disabledCategories.size()) {
      if (DEBUG_MODE) {
        WtMessageHandler.printToLogFile("config.getEnabledRuleIds().size() " + config.getEnabledRuleIds().size() 
            + "; allCaches.enabledRuleIds.size(): " + allCaches.enabledRuleIds.size() 
            + "\n config.getDisabledRuleIds().size(): " + config.getDisabledRuleIds().size()
            + "; allCaches.disabledRuleIds.size(): " + allCaches.disabledRuleIds.size() 
            + "\n config.getDisabledCategoryNames().size(): " + config.getDisabledCategoryNames().size()
            + "; allCaches.disabledCategories.size(): " + allCaches.disabledCategories.size()); 
      }
      return false;
    }
    for (String ruleId : config.getEnabledRuleIds()) {
      if (!allCaches.enabledRuleIds.contains(ruleId)) {
        return false;
      }
    }
    for (String category : config.getDisabledCategoryNames()) {
      if (!allCaches.disabledCategories.contains(category)) {
        return false;
      }
    }
    Set<String> disabledRuleIds = new HashSet<String>(config.getDisabledRuleIds());
    String langCode = WtOfficeTools.localeToString(mDocHandler.getLocale());
    for (String ruleId : WtDocumentsHandler.getDisabledRules(langCode)) {
      disabledRuleIds.add(ruleId);
    }
    if (disabledRuleIds.size() != allCaches.disabledRuleIds.size()) {
      return false;
    }
    for (String ruleId : disabledRuleIds) {
      if (!allCaches.disabledRuleIds.contains(ruleId)) {
        return false;
      }
    }
    Map<String, Set<String>> disabledRulesUI = new HashMap<String, Set<String>>();
    for (String lang : allCaches.disabledRulesUI.keySet()) {
      Set<String > ruleIDs = new HashSet<String>();
      for (String ruleID : allCaches.disabledRulesUI.get(lang)) {
        ruleIDs.add(ruleID);
      }
      disabledRulesUI.put(langCode, ruleIDs);
    }
    mDocHandler.setAllDisabledRules(disabledRulesUI);
    return true;
  }
  
  /**
   * get document cache
   */
  public WtDocumentCache getDocumentCache() {
    return allCaches.docCache;
  }
  
  /**
   * get paragraph caches (results for check of paragraphes)
   */
  public List<WtResultCache> getParagraphsCache() {
    return allCaches.paragraphsCache;
  }
  
  /**
   * get paragraph caches (results for check of paragraphes)
   */
  public WtResultCache getAiSuggestionCache() {
    return allCaches.aiSuggestionCache;
  }
  
  /**
   * get ignored matches
   */
  public WtIgnoredMatches getIgnoredMatches() {
    Map<Integer, Map<String, Set<Integer>>> ignoredMatches = new HashMap<>();
    for (int y : allCaches.ignoredMatches.keySet()) {
      Map<String, Set<Integer>> newIdMap = new HashMap<>();
      Map<String, Set<Integer>> idMap = new HashMap<>(allCaches.ignoredMatches.get(y));
      for (String id : idMap.keySet()) {
        Set<Integer> xSet = new HashSet<>(idMap.get(id));
        newIdMap.put(id, xSet);
      }
      ignoredMatches.put(y, newIdMap);
    }
    Map<Integer, List<LocaleEntry>> sLocales = new HashMap<>();
    WtMessageHandler.printToLogFile("CacheIO: getIgnoredMatches: spellLocales: size: " + allCaches.spellLocales.size());
    for (int y : allCaches.spellLocales.keySet()) {
      List<LocaleEntry> newEntryList = new ArrayList<>();
      List<LocaleSerialEntry> locEntries = new ArrayList<>(allCaches.spellLocales.get(y));
      WtMessageHandler.printToLogFile("CacheIO: getIgnoredMatches: spellLocales: size: " + locEntries.size() + " at y: " + y);
      for (LocaleSerialEntry entry : locEntries) {
        newEntryList.add(new LocaleEntry(entry.start, entry.length, entry.locale.toLocale(), entry.ruleId));
      }
      sLocales.put(y, newEntryList);
    }
    return new WtIgnoredMatches(ignoredMatches, sLocales);
  }
  
  /**
   * set all caches to null
   */
  public void resetAllCache() {
    allCaches = null;
  }
  
  /**
   * print debug information of caches to log file
   */
  private void printCacheInfo() {
    WtMessageHandler.printToLogFile("CacheIO: saveCaches:");
    WtMessageHandler.printToLogFile("Document Cache: Number of paragraphs: " + allCaches.docCache.size());
    WtMessageHandler.printToLogFile("Paragraph Cache(0): Number of paragraphs: " + allCaches.paragraphsCache.get(0).getNumberOfParas() 
        + ", Number of matches: " + allCaches.paragraphsCache.get(0).getNumberOfMatches());
    WtMessageHandler.printToLogFile("Paragraph Cache(1): Number of paragraphs: " + allCaches.paragraphsCache.get(1).getNumberOfParas() 
        + ", Number of matches: " + allCaches.paragraphsCache.get(1).getNumberOfMatches());
    for (int n = 0; n < allCaches.docCache.size(); n++) {
      WtMessageHandler.printToLogFile("allCaches.docCache.getFlatParagraphLocale(" + n + "): " 
            + (allCaches.docCache.getFlatParagraphLocale(n) == null ? "null" : WtOfficeTools.localeToString(allCaches.docCache.getFlatParagraphLocale(n))));
    }
    if (allCaches.paragraphsCache.get(0) == null) {
      WtMessageHandler.printToLogFile("paragraphsCache(0) == null");
    } else {
      if (allCaches.paragraphsCache.get(0).getNumberOfMatches() > 0) {
        for (int n = 0; n < allCaches.paragraphsCache.get(0).getNumberOfParas(); n++) {
          if (allCaches.paragraphsCache.get(0).getMatches(n, LoErrorType.BOTH) == null) {
            WtMessageHandler.printToLogFile("allCaches.sentencesCache.getMatches(" + n + ") == null");
          } else {
            if (allCaches.paragraphsCache.get(0).getMatches(n, LoErrorType.BOTH).length > 0) {
              WtMessageHandler.printToLogFile("Paragraph " + n + " sentence match[0]: " 
                  + "nStart = " + allCaches.paragraphsCache.get(0).getMatches(n, LoErrorType.BOTH)[0].nErrorStart 
                  + ", nLength = " + allCaches.paragraphsCache.get(0).getMatches(n, LoErrorType.BOTH)[0].nErrorLength
                  + ", errorID = " 
                  + (allCaches.paragraphsCache.get(0).getMatches(n, LoErrorType.BOTH)[0].aRuleIdentifier == null ? "null" 
                      : allCaches.paragraphsCache.get(0).getMatches(n, LoErrorType.BOTH)[0].aRuleIdentifier));
            }
          }
        }
      }
    }
  }

  class AllCaches implements Serializable {

    private static final long serialVersionUID = 6L;

    private final WtDocumentCache docCache;                               //  cache of paragraphs
    private final List<WtResultCache> paragraphsCache;                    //  Cache for matches of text rules
    private final WtResultCache aiSuggestionCache;                        //  Cache for AI results for other formulation of paragraph
    private final Map<String, List<String>> disabledRulesUI;              //  Disabled rules (per language)  
    private final List<String> disabledRuleIds;                           //  Disabled rules by configuration
    private final List<String> disabledCategories;                        //  Disabled categories
    private final List<String> enabledRuleIds;                            //  enabled rules by configuration
    private final Map<Integer, Map<String, Set<Integer>>> ignoredMatches; //  Map of matches (number of paragraph, number of character) that should be ignored after ignoreOnce was called
    private final Map<Integer, List<LocaleSerialEntry>> spellLocales;     //  Map of locales for ignored matches
    private final String ltVersion;                                       //  LT version
    
    AllCaches(WtDocumentCache docCache, List<WtResultCache> paragraphsCache, WtResultCache aiSuggestionCache,
        Map<String, Set<String>> disabledRulesUI, Set<String> disabledRuleIds, 
        Set<String> disabledCategories, Set<String> enabledRuleIds, WtIgnoredMatches ignoredMatches, String ltVersion) {
      this.docCache = docCache;
      this.paragraphsCache = paragraphsCache;
      this.aiSuggestionCache = aiSuggestionCache;
      this.disabledRulesUI = new HashMap<String, List<String>>();
      for (String langCode : disabledRulesUI.keySet()) {
        List <String >ruleIDs = new ArrayList<String>();
        for (String ruleID : disabledRulesUI.get(langCode)) {
          ruleIDs.add(ruleID);
        }
        this.disabledRulesUI.put(langCode, ruleIDs);
      }
      this.disabledRuleIds = new ArrayList<String>();
      for (String ruleID : disabledRuleIds) {
        this.disabledRuleIds.add(ruleID);
      }
      this.disabledCategories = new ArrayList<String>();
      for (String category : disabledCategories) {
        this.disabledCategories.add(category);
      }
      this.enabledRuleIds = new ArrayList<String>();
      for (String ruleID : enabledRuleIds) {
        this.enabledRuleIds.add(ruleID);
      }
      this.ltVersion = ltVersion;
      Map<Integer, Map<String, Set<Integer>>> clone = new HashMap<>();
      for (int y : ignoredMatches.getFullIMMap().keySet()) {
        Map<String, Set<Integer>> newIdMap = new HashMap<>();
        Map<String, Set<Integer>> idMap = new HashMap<>(ignoredMatches.get(y));
        for (String id : idMap.keySet()) {
          Set<Integer> xSet = new HashSet<>(idMap.get(id));
          newIdMap.put(id, xSet);
        }
        clone.put(y, newIdMap);
      }
      this.ignoredMatches = clone;
      Map<Integer, List<LocaleSerialEntry>> sLocales = new HashMap<>();
      for (int y : ignoredMatches.getFullSLMap().keySet()) {
        List<LocaleSerialEntry> newEntryList = new ArrayList<>();
        List<LocaleEntry> locEntries = new ArrayList<>(ignoredMatches.getLocaleEntries(y));
        WtMessageHandler.printToLogFile("CacheIO: AllCaches: spellLocales: size: " + locEntries.size() + " at y: " + y);
        for (LocaleEntry entry : locEntries) {
          newEntryList.add(new LocaleSerialEntry(entry.start, entry.length, entry.locale, entry.ruleId));
        }
        sLocales.put(y, newEntryList);
      }
      this.spellLocales = sLocales;
    }
    
  }
  
  public class LocaleSerialEntry  implements Serializable {

    private static final long serialVersionUID = 1L;

    int start;
    int length;
    SerialLocale locale;
    String ruleId;
    
    LocaleSerialEntry(int start, int length, Locale locale, String ruleId) {
      this.start = start;
      this.length = length;
      this.locale = new SerialLocale(locale);
      this.ruleId = new String(ruleId);
    }
  }

  /**
   * Class to to handle the cache files
   * cache files are stored in the LT configuration directory subdirectory 'cache'
   * the paths of documents are mapped to cache files and stored in the cache map
   */
  private class CacheFile implements Serializable {

    private static final long serialVersionUID = 1L;
    private CacheMap cacheMap;
    private File cacheMapFile;

    @SuppressWarnings("unused")
    CacheFile() {
      this(WtOfficeTools.getCacheDir());
    }

    CacheFile(File cacheDir) {
      cacheMapFile = new File(cacheDir, CACHEFILE_MAP);
      if (cacheMapFile != null) {
        if (cacheMapFile.exists() && !cacheMapFile.isDirectory()) {
          if (read()) {
            return;
          }
        }
        cacheMap = new CacheMap();
        if (DEBUG_MODE) {
          WtMessageHandler.printToLogFile("CacheIO: CacheFile: create cacheMap file");
        }
        write(cacheMap);
      }
    }

    /**
     * read the cache map from file
     */
    public boolean read() {
      try {
        FileInputStream fileIn = new FileInputStream(cacheMapFile);
        ObjectInputStream in = new ObjectInputStream(fileIn);
        cacheMap = (CacheMap) in.readObject();
        if (DEBUG_MODE) {
          WtMessageHandler.printToLogFile("CacheIO: CacheFile: read cacheMap file: size=" + cacheMap.size());
        }
        in.close();
        fileIn.close();
        return true;
      } catch (Throwable t) {
        WtMessageHandler.printException(t);     // all Exceptions thrown are printed
        return false;
      }
    }

    /**
     * write the cache map to file
     */
    public void write(CacheMap cacheMap) {
      try {
        FileOutputStream fileOut = new FileOutputStream(cacheMapFile);
        ObjectOutputStream out = new ObjectOutputStream(fileOut);
        if (DEBUG_MODE) {
          WtMessageHandler.printToLogFile("CacheIO: CacheFile: write cacheMap file: size=" + cacheMap.size());
        }
        out.writeObject(cacheMap);
        out.close();
        fileOut.close();
      } catch (Throwable t) {
        WtMessageHandler.printException(t);     // all Exceptions thrown are printed
      }
    }
    
    /**
     * get the cache file name for a given document path
     * if create == true: create a cache file if it not exists 
     */
    public String getCacheFileName(String docPath, boolean create) {
      if (cacheMap == null) {
        return null;
      }
      int orgSize = cacheMap.size();
      String cacheFileName = cacheMap.getOrCreateCacheFile(docPath, create);
      if (cacheMap.size() != orgSize) {
        write(cacheMap);
      }
      return cacheFileName;
    }
    
   /**
    * remove unused files from cache directory
    */
    public void cleanUp(String curCacheFile) {
      CacheCleanUp cacheCleanUp = new CacheCleanUp(cacheMap, curCacheFile);
      cacheCleanUp.start();
    }
    
    /**
     * Class to create and handle the cache map
     * the clean up process is run in a separate parallel thread
     */
    private class CacheMap implements Serializable {
      private static final long serialVersionUID = 1L;
      private Map<String, String> cacheNames;     //  contains the mapping from document paths to cache file names
      
      CacheMap() {
        cacheNames = new HashMap<>();
      }
      
      CacheMap(CacheMap in) {
        cacheNames = new HashMap<>();
        cacheNames.putAll(in.getCacheNames());
      }

      /**
       * get the cache map that contains the mapping from document paths to cache file names 
       */
      private Map<String, String> getCacheNames() {
        return cacheNames;
      }

      /**
       * get the size of the cache map
       */
      public int size() {
        return cacheNames.keySet().size();
      }
      
      /**
       * return true if the map contains the cache file name
       */
      public boolean containsValue(String value) {
        return cacheNames.containsValue(value);
      }
      
      /**
       * get all document paths contained in the map
       */
      public Set<String> keySet() {
        return cacheNames.keySet();
      }
      
      /**
       * get the cache file name from a document path
       */
      public String get(String key) {
        return cacheNames.get(key);
      }
      
      /**
       * remove a document paths from the map (inclusive the mapped cache file name)
       */
      public String remove(String key) {
        return cacheNames.remove(key);
      }
      
      /**
       * get the cache file name for a document paths
       * if create == true:  create the file if not exist
       * if create == false: return null if not exist
       */
      public String getOrCreateCacheFile(String docPath, boolean create) {
        if (DEBUG_MODE) {
          WtMessageHandler.printToLogFile("CacheIO: getOrCreateCacheFile: docPath=" + docPath);
          for (String file : cacheNames.keySet()) {
            WtMessageHandler.printToLogFile("cacheNames: docPath=" + file + ", cache=" + cacheNames.get(file));
          }
        }
        if (cacheNames.containsKey(docPath)) {
          return cacheNames.get(docPath);
        }
        if (!create) {
          return null;
        }
        int i = 1;
        String cacheName = CACHEFILE_PREFIX + i + "." + CACHEFILE_EXTENSION;
        while (cacheNames.containsValue(cacheName)) {
          i++;
          cacheName = CACHEFILE_PREFIX + i + "." + CACHEFILE_EXTENSION;
        }
        cacheNames.put(docPath, cacheName);
        return cacheName;
      }
    }
    
    /**
     * class to clean up cache
     * delete cache files for not existent document paths
     * remove not existent document paths and cache files from map
     * the clean up process is ran in a separate thread 
     */
    private class CacheCleanUp extends Thread implements Serializable {
      private static final long serialVersionUID = 1L;
      private CacheMap cacheMap;
      private String currentFile;
      
      CacheCleanUp(CacheMap in, String curFile) {
        cacheMap = new CacheMap(in);
        currentFile = curFile;
      }
      
      /**
       * run clean up process
       */
      @Override
      public void run() {
        try {
          long systemTime = System.currentTimeMillis();
          boolean mapChanged = false;
          File cacheDir = WtOfficeTools.getCacheDir();
          List<String> mapedDocs = new ArrayList<String>();
          for (String doc : cacheMap.keySet()) {
            mapedDocs.add(doc);
          }
          for (String doc : mapedDocs) {
            File docFile = new File(doc);
            String cacheFileName = cacheMap.get(doc);
            File cacheFile = new File(cacheDir, cacheFileName);
            if (DEBUG_MODE) {
              WtMessageHandler.printToLogFile("CacheIO: CacheCleanUp: CacheMap: docPath=" + doc + ", docFile exist: " + (docFile == null ? "null" : docFile.exists()) + 
                  ", cacheFile exist: " + (cacheFile == null ? "null" : cacheFile.exists()));
            }
            if (docFile == null || !docFile.exists() || cacheFile == null || !cacheFile.exists() 
                || (systemTime - cacheFile.lastModified() > MAX_CACHE_TIME && !cacheFileName.equals(currentFile))) {
              cacheMap.remove(doc);
              mapChanged = true;
              WtMessageHandler.printToLogFile("CacheIO: CacheCleanUp: Remove Path from CacheMap: " + doc);
              if (cacheFile != null && cacheFile.exists()) {
                cacheFile.delete();
                WtMessageHandler.printToLogFile("CacheIO: CacheCleanUp: Delete cache file: " + cacheFile.getAbsolutePath());
              }
            }
          }
          if (mapChanged) {
            if (DEBUG_MODE) {
              WtMessageHandler.printToLogFile("CacheIO: CacheCleanUp: Write CacheMap");
            }
            write(cacheMap);
          }
          File[] cacheFiles = cacheDir.listFiles();
          if (cacheFiles != null) {
            for (File cacheFile : cacheFiles) {
              if (!cacheMap.containsValue(cacheFile.getName()) && !cacheFile.getName().equals(CACHEFILE_MAP)
                  && !cacheFile.getName().equals(SPELL_CACHEFILE)) {
                cacheFile.delete();
                WtMessageHandler.printToLogFile("Delete cache file: " + cacheFile.getAbsolutePath());
              }
            }
          }
        } catch (Throwable t) {
          WtMessageHandler.printException(t);     // all Exceptions thrown are printed
        }
      }
    }
  }

  public class SpellCache implements Serializable {
    private static final long serialVersionUID = 1L;
    private final Map<String, List<String>> lastWrongWords = new HashMap<>();
    private final Map<String, List<String[]>> lastSuggestions = new HashMap<>();
    private String version = WtVersionInfo.ltVersion();
    private int statVersion = 1;
    
    private boolean putAll (SpellCache sc) {
      version = sc.version;
      if (sc.statVersion != statVersion || !version.equals(WtVersionInfo.ltVersion())) {
        return false;
      }
      lastWrongWords.clear();
      lastWrongWords.putAll(sc.lastWrongWords);
      lastSuggestions.clear();
      lastSuggestions.putAll(sc.lastSuggestions);
      return true;
    }
    
    public void write(Map<String, List<String>> lastWrongWords, Map<String, List<String[]>> lastSuggestions) {
      this.lastWrongWords.clear();
      this.lastSuggestions.clear();
      this.lastWrongWords.putAll(lastWrongWords);
      this.lastSuggestions.putAll(lastSuggestions);
      try {
        File cacheDir = WtOfficeTools.getCacheDir();
        File cacheFilePath = new File(cacheDir, SPELL_CACHEFILE);
        String cachePath = cacheFilePath.getAbsolutePath();
        if (cachePath != null) {
          GZIPOutputStream fileOut = new GZIPOutputStream(new FileOutputStream(cachePath));
          ObjectOutputStream out = new ObjectOutputStream(fileOut);
          out.writeObject(this);
          out.close();
          fileOut.close();
          WtMessageHandler.printToLogFile("Spell Cache saved to: " + cachePath);
        }
      } catch (Throwable t) {
        WtMessageHandler.printException(t);     // all Exceptions thrown are printed
      }
    }
    
    public boolean read() {
      File cacheDir = WtOfficeTools.getCacheDir();
      File cacheFilePath = new File(cacheDir, SPELL_CACHEFILE);
      String cachePath = cacheFilePath.getAbsolutePath();
      if (cachePath == null) {
        return false;
      }
      try {
        File file = new File(cachePath);
        if (file.exists() && !file.isDirectory()) {
          GZIPInputStream fileIn = new GZIPInputStream(new FileInputStream(file));
          ObjectInputStream in = new ObjectInputStream(fileIn);
          boolean out = putAll((SpellCache) in.readObject());
          in.close();
          fileIn.close();
          WtMessageHandler.printToLogFile("Spell Cache read from: " + cachePath);
          if (out) {
            return true;
          } else {
            WtMessageHandler.printToLogFile("Version has changed: Spell Cache rejected (Cache Version: " 
                  + version + ", actual LT Version: " + WtVersionInfo.ltVersion() + ")");
            return false;
          }
        }
      } catch (InvalidClassException e) {
        WtMessageHandler.printToLogFile("Old cache Version: Spell Cache not read");
        return false;
      } catch (Throwable t) {
        WtMessageHandler.printException(t);     // all Exceptions thrown are printed
      }
      return false;
    }
    
    public Map<String, List<String>> getWrongWords() {
      return lastWrongWords;
    }
    
    public Map<String, List<String[]>> getSuggestions() {
      return lastSuggestions;
    }
    
  }

  
}
