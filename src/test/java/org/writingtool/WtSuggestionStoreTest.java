package org.writingtool;

import static org.junit.Assert.assertArrayEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;

import com.sun.star.lang.Locale;

public class WtSuggestionStoreTest {

  @Test
  public void testSuggestionsStayInMostRecentOrder() throws Throwable {
    Locale locale = new Locale("en", "US", "");
    WtSuggestionStore store = new WtSuggestionStore(2);

    store.addSuggestions("alpha", locale, new String[] { "alpha-1" });
    store.addSuggestions("beta", locale, new String[] { "beta-1" });

    assertArrayEquals(new String[] { "alpha-1" }, store.getSuggestions("alpha", locale));

    store.addSuggestions("gamma", locale, new String[] { "gamma-1" });

    assertTrue(store.hasWord("alpha", locale));
    assertFalse(store.hasWord("beta", locale));
    assertTrue(store.hasWord("gamma", locale));
  }

  @Test
  public void testSetFullStoreUsesProvidedOrder() throws Throwable {
    Locale locale = new Locale("en", "US", "");
    WtSuggestionStore store = new WtSuggestionStore(3);
    Map<String, List<String>> words = new HashMap<>();
    Map<String, List<String[]>> suggestions = new HashMap<>();

    words.put("en-US", Arrays.asList("first", "second"));
    suggestions.put("en-US", Arrays.asList(new String[] { "first-1" }, new String[] { "second-1" }));

    store.setFullStore(words, suggestions);

    assertArrayEquals(new String[] { "first-1" }, store.getSuggestions("first", locale));
    assertArrayEquals(new String[] { "second-1" }, store.getSuggestions("second", locale));
  }
}
