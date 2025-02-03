package org.writingtool;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Pattern;

public class WtQuotesDetection {
  
  private static final List<String> startSymbols = Arrays.asList("“", "„", "»", "«", "\"");
  private static final List<String> endSymbols   = Arrays.asList("”", "“", "«", "»", "\"");
  private static final Pattern PUNCTUATION = Pattern.compile("[\\p{Punct}…–—&&[^\"'_]]");
  private static final Pattern PUNCT_MARKS = Pattern.compile("[\\?\\.!,]");
  
  private List<Integer> openingQuotes;
  private List<Integer> closingQuotes;
  
  private String openQuote = new String();

  /**
   * Defines exceptions from quotes
   */
  private boolean isNotQuote (String txt, int i, int j) {
    if ((i == 0 || Character.isWhitespace(txt.charAt(i - 1)))
        && (i >= txt.length() - 1 || Character.isWhitespace(txt.charAt(i + 1)))) {
      return true;
    }
    if (endSymbols.get(j).equals(startSymbols.get(j))) {
      if (i > 0 && i < txt.length() - 1 
          && !Character.isWhitespace(txt.charAt(i - 1))
          && !Character.isWhitespace(txt.charAt(i + 1))) {
        String pChar = txt.substring(i - 1, i);
        String nChar = txt.substring(i + 1, i + 2);
        if(PUNCTUATION.matcher(pChar).matches()
          && !".".equals(nChar)
          && PUNCTUATION.matcher(nChar).matches()) {
          return true;
        }
      }
    }
    return false;
  }
  
  /**
   * defines an opening quote
   */
  private boolean isOpeningQuote(String txt, int i) {
    String tChar = txt.substring(i, i + 1);
    for (int j = 0; j < startSymbols.size(); j++) {
      if (startSymbols.get(j).equals(tChar)) {
        if (isNotQuote (txt, i, j)) {
          return false;
        }
        if (endSymbols.contains(startSymbols.get(j))) {
          return (i == 0
              || Character.isWhitespace(txt.charAt(i - 1))
              || (i < txt.length() - 1 && !Character.isWhitespace(txt.charAt(i + 1))
                  && ((!PUNCT_MARKS.matcher(txt.substring(i + 1, i + 2)).matches()
                      && PUNCTUATION.matcher(txt.substring(i - 1, i)).matches())
                      || (txt.charAt(i + 1) == '-'))));
        }
        return true;
      }
    }
    return false;
  }

  /**
   * defines an closing quote
   */
  private boolean isClosingQuote(String txt, int i) {
    String tChar = txt.substring(i, i + 1);
    for (int j = 0; j < endSymbols.size(); j++) {
      if (endSymbols.get(j).equals(tChar)) {
        if (isNotQuote (txt, i, j) && !openQuote.equals(startSymbols.get(j))) {
          return false;
        }
        return true;
      }
    }
    return false;
  }
  
  /**
   * analyse one paragraph
   */
  private boolean analyzeOneParagraph(String txt, boolean isQuoteBefore) {
    openingQuotes = new ArrayList<>();
    closingQuotes = new ArrayList<>();
    if (isQuoteBefore) {
      openingQuotes.add(-1);
    }
    for (int i = 0; i < txt.length(); i++) {
      if (isOpeningQuote(txt, i)) {
        openingQuotes.add(i);
        openQuote = new String(txt.substring(i, i + 1));
        isQuoteBefore = true;
      } else if (isClosingQuote(txt, i)) {
        closingQuotes.add(i);
        openQuote = new String();
        isQuoteBefore = false;
      }
    }
    return isQuoteBefore;
  }
  
  /**
   * analyse all paragraphs
   * NOTE: only text paragraphs will be analyzed (no footnotes, tables, etc.)
   * NOTE: Quotes lists have to initialized
   */
  public void analyzeTextParagraphs(List<String> paragraphs, List<List<Integer>> oQuotes, List<List<Integer>> cQuotes) {
    oQuotes.clear();
    cQuotes.clear();
    boolean isQuoteBefore = false;
    for (int i = 0; i < paragraphs.size(); i++) {
      isQuoteBefore = analyzeOneParagraph(paragraphs.get(i), isQuoteBefore);
      oQuotes.add(openingQuotes);
      cQuotes.add(closingQuotes);
    }
  }


}
