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

import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.locks.ReentrantReadWriteLock;

import org.writingtool.tools.WtOfficeTools.LoErrorType;

import com.sun.star.linguistic2.SingleProofreadingError;
import com.sun.star.text.TextMarkupType;

/**
 * Class for storing and handle the LT results prepared to use in LO/OO
 *
 * @author Fred Kruse
 * @since 1.0
 */
public class WtResultCache implements Serializable {

  private static final long serialVersionUID = 2L;
  private final Map<Integer, SerialCacheEntry> entries = new HashMap<Integer, SerialCacheEntry>();
  
  private ReentrantReadWriteLock rwLock = new ReentrantReadWriteLock();
  
  public WtResultCache() {
    this(null);
  }

  public WtResultCache(WtResultCache cache) {
    replace(cache);
  }

  /**
   * Get cache entry map
   */
  private Map<Integer, SerialCacheEntry> getMap() {
    rwLock.readLock().lock();
    try {
      return new HashMap<Integer, SerialCacheEntry>(entries);
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * Replace the cache content
   */
  public void replace(WtResultCache cache) {
    rwLock.writeLock().lock();
    try {
      entries.clear();
      if (cache != null && !cache.entries.isEmpty()) {
        entries.putAll(cache.getMap());
      }
    } finally {
      rwLock.writeLock().unlock();
    }
  }

  /**
   * Remove all cache entries for a paragraph
   */
  public void remove(int numberOfParagraph) {
    rwLock.writeLock().lock();
    try {
      entries.remove(numberOfParagraph);
    } finally {
      rwLock.writeLock().unlock();
    }
  }

  /**
   * Remove all cache entries between firstParagraph and lastParagraph
   */
  public void removeRange(int firstParagraph, int lastParagraph) {
    rwLock.writeLock().lock();
    try {
      for (int i = firstParagraph; i <= lastParagraph; i++) {
        entries.remove(i);
      }
    } finally {
      rwLock.writeLock().unlock();
    }
  }

  /**
   * Remove all cache entries between firstPara (included) and lastPara (excluded)
   * shift all numberOfParagraph by 'shift'
   */
  public void removeAndShift(int fromParagraph, int toParagraph, int oldSize, int newSize) {
    int shift = newSize - oldSize;
    if (fromParagraph < 0 && toParagraph >= newSize) {
      return;
    }
    rwLock.writeLock().lock();
    try {
      Map<Integer, SerialCacheEntry> tmpEntries = new HashMap<Integer, SerialCacheEntry>(entries);
      entries.clear();
      if (shift < 0) {   // new size < old size
        for (int i : tmpEntries.keySet()) {
          if (i < fromParagraph) {
            entries.put(i, tmpEntries.get(i));
          } else if (i >= toParagraph - shift) {
              entries.put(i + shift, tmpEntries.get(i));
          }
        }
      } else {
        for (int i : tmpEntries.keySet()) {
          if (i < fromParagraph) {
            entries.put(i, tmpEntries.get(i));
          } else if (i >= toParagraph) {
              entries.put(i + shift, tmpEntries.get(i));
          }
        }
      }
    } finally {
      rwLock.writeLock().unlock();
    }
  }

  /**
   * add or replace a cache entry
   */
  public void put(int numberOfParagraph, List<Integer> nextSentencePositions, WtProofreadingError[] errorArray) {
    rwLock.writeLock().lock();
    try {
      entries.put(numberOfParagraph, new SerialCacheEntry(nextSentencePositions, errorArray));
    } finally {
      rwLock.writeLock().unlock();
    }
  }

  /**
   * add or replace a cache entry for paragraph
   */
  public void put(int numberOfParagraph, WtProofreadingError[] errorArray) {
    rwLock.writeLock().lock();
    try {
      entries.put(numberOfParagraph, new SerialCacheEntry(null, errorArray));
    } finally {
      rwLock.writeLock().unlock();
    }
  }

  /**
   * add proof reading errors to a cache entry for paragraph
   */
  public void add(int numberOfParagraph, WtProofreadingError[] errorArray) {
    rwLock.writeLock().lock();
    try {
      SerialCacheEntry cacheEntry = entries.get(numberOfParagraph);
      cacheEntry.addErrorArray(errorArray);
      entries.put(numberOfParagraph, cacheEntry);
    } finally {
      rwLock.writeLock().unlock();
    }
  }

  /**
   * Remove all cache entries
   */
  public void removeAll() {
    rwLock.writeLock().lock();
    try {
      entries.clear();
    } finally {
      rwLock.writeLock().unlock();
    }
  }

  /**
   * Size of cache (size of entries)
   */
  public int size() {
    rwLock.readLock().lock();
    try {
      return entries.size();
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get cache entry of paragraph
   */
  public int getNumberofNotNullEntries() {
    rwLock.readLock().lock();
    try {
      int num = 0;
      for (int n : entries.keySet()) {
        if (entries.get(n) != null) {
          num++;
        }
      }
      return num;
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get cache entry of paragraph
   */
  public int getNumberofErrors() {
    rwLock.readLock().lock();
    try {
      int num = 0;
      for (int n : entries.keySet()) {
        if (entries.get(n) != null) {
          num += entries.get(n).errorArray.length;
        }
      }
      return num;
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get cache entry of paragraph
   */
  public CacheEntry getCacheEntry(int numberOfParagraph) {
    rwLock.readLock().lock();
    try {
      SerialCacheEntry entry = entries.get(numberOfParagraph);
      return entry == null ? null : new CacheEntry(entry);
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get cache entry of paragraph without read lock 
   */
  public CacheEntry getUnsafeCacheEntry(int numberOfParagraph) {
    SerialCacheEntry entry = entries.get(numberOfParagraph);
    return entry == null ? null : new CacheEntry(entry);
  }

  /**
   * cache has an Error or has reached a limit of entries
   */
  public boolean hasAnError(int limit) {
    rwLock.readLock().lock();
    try {
      if (entries.size() >= limit) {
        return true;
      }
      Set<Integer> paras = new HashSet<>(entries.keySet());
      for (int n : paras) {
        if (entries.get(n).errorArray.length > 0) {
          return true;
        }
      }
      return false;
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get cache entry of paragraph
   */
  public SerialCacheEntry getSerialCacheEntry(int numberOfParagraph) {
    rwLock.readLock().lock();
    try {
      return entries.get(numberOfParagraph);
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get Proofreading errors of on paragraph from cache
   * get an error array, if a prargrapherray exists
   */
  public WtProofreadingError[] getSafeMatches(int numberOfParagraph) {
    rwLock.readLock().lock();
    try {
      SerialCacheEntry entry = entries.get(numberOfParagraph);
      if (entry == null) {
        return null;
      }
      if (entry.errorArray == null) {
        return (new WtProofreadingError[0]);
      }
      return entry.getWtErrorArray();
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get Proofreading errors of on paragraph from cache
   */
  public WtProofreadingError[] getMatches(int numberOfParagraph, LoErrorType errType) {
    rwLock.readLock().lock();
    try {
      SerialCacheEntry entry = entries.get(numberOfParagraph);
      if (entry == null) {
        return null;
      }
      WtProofreadingError[] errorArray = entry.getWtErrorArray();
      if (errType == LoErrorType.BOTH || errorArray == null || errorArray.length == 0) {
        return errorArray;
      }
      List<WtProofreadingError> errorList = new ArrayList<>();
      for (WtProofreadingError eArray : errorArray) {
        if ((errType == LoErrorType.GRAMMAR && eArray.nErrorType == TextMarkupType.PROOFREADING)
            || (errType == LoErrorType.SPELL && eArray.nErrorType == TextMarkupType.SPELLCHECK)) {
          errorList.add(eArray);
        }
      }
      return errorList.toArray(new WtProofreadingError[errorList.size()]);
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get start sentence position from cache
   */
  public int getStartSentencePosition(int numberOfParagraph, int sentencePosition) {
    rwLock.readLock().lock();
    try {
      SerialCacheEntry entry = entries.get(numberOfParagraph);
      if (entry == null) {
        return 0;
      }
      List<Integer> nextSentencePositions = entry.nextSentencePositions;
      if (nextSentencePositions == null || nextSentencePositions.size() < 2) {
        return 0;
      }
      int startPosition = 0;
      for (int position : nextSentencePositions) {
        if (position >= sentencePosition) {
          return position == sentencePosition ? position : startPosition;
        }
        startPosition = position;
      }
      return nextSentencePositions.get(nextSentencePositions.size() - 2);
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get next sentence position from cache
   */
  public int getNextSentencePosition(int numberOfParagraph, int sentencePosition) {
    rwLock.readLock().lock();
    try {
      SerialCacheEntry entry = entries.get(numberOfParagraph);
      if (entry == null) {
        return 0;
      }
      List<Integer> nextSentencePositions = entry.nextSentencePositions;
      if (nextSentencePositions == null || nextSentencePositions.size() == 0) {
        return 0;
      }
      for (int position : nextSentencePositions) {
        if (position > sentencePosition) {
          return position;
        }
      }
      return nextSentencePositions.get(nextSentencePositions.size() - 1);
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get Proofreading errors of sentence out of paragraph matches from cache
   */
  public WtProofreadingError[] getFromPara(int numberOfParagraph,
                                        int startOfSentencePosition, int endOfSentencePosition, LoErrorType errType) {
    rwLock.readLock().lock();
    try {
      SerialCacheEntry entry = entries.get(numberOfParagraph);
      if (entry == null) {
        return null;
      }
      List<WtProofreadingError> errorList = new ArrayList<>();
      for (WtProofreadingError eArray : entry.getWtErrorArray()) {
        if ((errType == LoErrorType.BOTH 
            || (errType == LoErrorType.GRAMMAR && eArray.nErrorType == TextMarkupType.PROOFREADING)
            || (errType == LoErrorType.SPELL && eArray.nErrorType == TextMarkupType.SPELLCHECK))
            && eArray.nErrorStart >= startOfSentencePosition && eArray.nErrorStart < endOfSentencePosition) {
          errorList.add(eArray);
        }
      }
      return errorList.toArray(new WtProofreadingError[0]);
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * Compares to Entries
   * true if the both entries are NOT identically
   */
  public static boolean areDifferentEntries(SerialCacheEntry newEntries, SerialCacheEntry oldEntries) {
    if (newEntries == null || oldEntries == null) {
      return true;
    }
    WtProofreadingError[] oldErrorArray = oldEntries.getWtErrorArray();
    WtProofreadingError[] newErrorArray = newEntries.getWtErrorArray();
    if (oldErrorArray == null || newErrorArray == null || oldErrorArray.length != newErrorArray.length) {
      return true;
    }
    for (WtProofreadingError nError : newErrorArray) {
      if (nError.nErrorType != TextMarkupType.SPELLCHECK) {
        boolean found = false;
        for (WtProofreadingError oError : oldErrorArray) {
            if (nError.nErrorType != TextMarkupType.SPELLCHECK
                && nError.nErrorStart == oError.nErrorStart && nError.nErrorLength == oError.nErrorLength
                    && nError.aRuleIdentifier.equals(oError.aRuleIdentifier)) {
              found = true;
              break;
            }
          }
        if (!found) {
          return true;
        }
      }
    }
    return false;
  }

  /**
   * true if entry has no error
   */
  public static boolean isEmptyEntry(SerialCacheEntry entry) {
    if (entry == null || entry.errorArray == null || entry.errorArray.length == 0) {
      return true;
    }
    return false;
  }

  /**
   * Compares a paragraph cache with another cache.
   * Gives back a list of entries for every paragraph: true if the both entries are identically
   */
  public List<Integer> differenceInCaches(WtResultCache oldCache) {
    rwLock.readLock().lock();
    try {
      List<Integer> differentParas = new ArrayList<>();
      SerialCacheEntry oEntry;
      SerialCacheEntry nEntry;
      boolean isDifferent = true;
      Set<Integer> entrySet = new HashSet<>(entries.keySet());
      for (int nPara : entrySet) {
        if (oldCache != null) {
          nEntry = entries.get(nPara);
          oEntry = oldCache.entries.get(nPara);
          isDifferent = areDifferentEntries(nEntry, oEntry);
        }
        if (isDifferent) {
          differentParas.add(nPara);
        }
      }
      return differentParas;
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * Remove a special Proofreading error from cache
   * Returns all changed paragraphs as list
   */
  public List<Integer> removeRuleError(String ruleId) {
    rwLock.writeLock().lock();
    try {
      List<Integer> changed = new ArrayList<>();
      Set<Integer> keySet = entries.keySet();
      for (int n : keySet) {
        SerialCacheEntry entry = entries.get(n);
        WtProofreadingError[] eArray = entry.getWtErrorArray();
        int nErr = 0;
        for (WtProofreadingError sError : eArray) {
          if (sError.aRuleIdentifier.equals(ruleId)) {
            nErr++;
          }
        }
        if (nErr > 0) {
          changed.add(n);
          WtProofreadingError[] newArray = new WtProofreadingError[eArray.length - nErr];
          for (int i = 0, j = 0; i < eArray.length && j < newArray.length; i++) {
            if (!eArray[i].aRuleIdentifier.equals(ruleId)) {
              newArray[j] = eArray[i];
              j++;
            }
          }
          entries.put(n, new SerialCacheEntry(entry.nextSentencePositions, newArray));
        }
      }
      return changed;
    } finally {
      rwLock.writeLock().unlock();
    }
  }

  /**
   * get number of paragraphs stored in cache
   */
  public int getNumberOfParas() {
    rwLock.readLock().lock();
    try {
      return entries.size();
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get number of entries
   */
  public int getNumberOfEntries() {
    rwLock.readLock().lock();
    try {
      return entries.size();
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get number of matches
   */
  public int getNumberOfMatches() {
    rwLock.readLock().lock();
    try {
      int number = 0;
      for (int n : entries.keySet()) {
        number += entries.get(n).errorArray.length;
      }
      return number;
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get all errors from a position within a paragraph as List
   */
  public List<WtProofreadingError> getErrorsAtPosition(int numPara, int numChar) {
    rwLock.readLock().lock();
    List<WtProofreadingError> errors = new ArrayList<>();
    try {
      SerialCacheEntry entry = entries.get(numPara);
      if (entry == null) {
        return null;
      }
      for (WtProofreadingError err : entry.getWtErrorArray()) {
        if (numChar >= err.nErrorStart && numChar < err.nErrorStart + err.nErrorLength) {
          errors.add(err);
        }
      }
      return errors;
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * Class of serializable cache entries
   */
  public static class CacheEntry {
    public SingleProofreadingError[] errorArray;
    public List<Integer> nextSentencePositions = null;

    CacheEntry(SerialCacheEntry entry) {
      if (entry.nextSentencePositions != null) {
        this.nextSentencePositions = new ArrayList<Integer>(entry.nextSentencePositions);
      }
      this.errorArray = new SingleProofreadingError[entry.errorArray.length];
      for (int i = 0; i < entry.errorArray.length; i++) {
        this.errorArray[i] = entry.errorArray[i].toSingleProofreadingError();
      }
    }

    /**
     * Get an SingleProofreadingError array for one entry
     */
    public SingleProofreadingError[] getErrorArray() {
      return errorArray;
    }
  }
    
  /**
   * Class of serializable cache entries
   */
  class SerialCacheEntry implements Serializable {
    private static final long serialVersionUID = 2L;
    WtProofreadingError[] errorArray;
    List<Integer> nextSentencePositions = null;

    SerialCacheEntry(List<Integer> nextSentencePositions, WtProofreadingError[] sErrorArray) {
      if (nextSentencePositions != null) {
        this.nextSentencePositions = new ArrayList<Integer>(nextSentencePositions);
      }
      this.errorArray = new WtProofreadingError[sErrorArray.length];
      for (int i = 0; i < sErrorArray.length; i++) {
        this.errorArray[i] = new WtProofreadingError(sErrorArray[i]);
      }
    }
    
    /**
     * Get an WtProofreadingError array for one entry
     */
    WtProofreadingError[] getWtErrorArray() {
      return errorArray;
    }
    
    /**
     * Get an SingleProofreadingError array for one entry
     */
    SingleProofreadingError[] getSingleErrorArray() {
      SingleProofreadingError[] eArray = new SingleProofreadingError[errorArray.length];
      for (int i = 0; i < errorArray.length; i++) {
        eArray[i] = errorArray[i].toSingleProofreadingError();
      }
      return eArray;
    }
    
    /**
     * Get an SingleProofreadingError array for one entry
     */
    int getErrorSize() {
      return errorArray.length;
    }
    
    /**
     * Add an SingleProofreadingError array to an existing one
     */
    void addErrorArray(WtProofreadingError[] errors) {
      if (errors == null || errors.length == 0) {
        return;
      }
      WtProofreadingError newErrorArray[] = new WtProofreadingError[errorArray.length + errors.length];
      for (int i = 0; i < errorArray.length; i++) {
        newErrorArray[i] = errorArray[i];
      }
      for (int i = 0; i < errors.length; i++) {
        newErrorArray[errorArray.length + i] = new WtProofreadingError(errors[i]);
      }
    }
  }
  
}
