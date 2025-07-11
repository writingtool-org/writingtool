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
package org.writingtool.config;

import java.awt.Color;
import java.awt.Font;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.EnumMap;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.regex.Pattern;
import java.util.Properties;
import java.util.Set;

import org.apache.commons.lang3.StringUtils;
import org.jetbrains.annotations.Nullable;
import org.languagetool.Language;
import org.languagetool.Languages;
import org.languagetool.rules.ITSIssueType;
import org.languagetool.rules.Rule;
import org.languagetool.rules.RuleOption;
import org.writingtool.tools.WtGeneralTools;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtVersionInfo;

/**
 * Configuration like list of disabled rule IDs, server mode etc.
 * Configuration is loaded from and stored to a properties file.
 *
 * @author Fred Kruse
 */
public class WtConfiguration {
  
  public final static short UNDERLINE_WAVE = 10;
  public final static short UNDERLINE_BOLDWAVE = 18;
  public final static short UNDERLINE_BOLD = 12;
  public final static short UNDERLINE_DASH = 5;
  
  public final static int CHECK_DIRECT_SPEECH_YES = 0;
  public final static int CHECK_DIRECT_SPEECH_NO_STYLE = 1;
  public final static int CHECK_DIRECT_SPEECH_NO = 2;
  
  static final int DEFAULT_SERVER_PORT = 8081;  // should be HTTPServerConfig.DEFAULT_PORT but we don't have that dependency
  static final int DEFAULT_NUM_CHECK_PARAS = -2;  //  default number of parameters to be checked by TextLevelRules in LO/OO 
  static final int FONT_STYLE_INVALID = -1;
  static final int FONT_SIZE_INVALID = -1;
  static final int DEFAULT_COLOR_SELECTION = 0;
  static final int DEFAULT_CHECK_DIRECT_SPEECH = CHECK_DIRECT_SPEECH_YES;
  static final int DEFAULT_THEME_SELECTION = WtGeneralTools.THEME_SYSTEM;
  static final boolean DEFAULT_DO_RESET = false;
  static final boolean DEFAULT_MULTI_THREAD = false;
  static final boolean DEFAULT_NO_BACKGROUND_CHECK = false;
  static final boolean DEFAULT_USE_QUEUE = false;
  static final boolean DEFAULT_USE_DOC_LANGUAGE = true;
  static final boolean DEFAULT_DO_REMOTE_CHECK = false;
  static final boolean DEFAULT_USE_OTHER_SERVER = false;
  static final boolean DEFAULT_IS_PREMIUM = false;
  static final boolean DEFAULT_MARK_SINGLE_CHAR_BOLD = false;
  static final boolean DEFAULT_USE_LT_SPELL_CHECKER = true;
  static final boolean DEFAULT_USE_LONG_MESSAGES = false;
  static final boolean DEFAULT_NO_SYNONYMS_AS_SUGGESTIONS = false;
  static final boolean DEFAULT_INCLUDE_TRACKED_CHANGES = false;
  static final boolean DEFAULT_ENABLE_TMP_OFF_RULES = false;
  static final boolean DEFAULT_ENABLE_GOAL_SPECIFIC_RULES = false;
  static final boolean DEFAULT_FILTER_OVERLAPPING_MATCHES = true;
  static final boolean DEFAULT_SAVE_LO_CACHE = true;
  static final boolean DEFAULT_USE_AI_SUPPORT = false;
  static final boolean DEFAULT_USE_AI_IMG_SUPPORT = false;
  static final boolean DEFAULT_USE_AI_TTS_SUPPORT = false;
  static final boolean DEFAULT_AI_AUTO_CORRECT = true;
  static final boolean DEFAULT_AI_AUTO_SUGGESTION = false;
  static final int DEFAULT_AI_SHOW_STYLISTIC_CHANGES = 0;
  
  static final String DEFAULT_AI_MODEL = "gpt-4";
  static final String DEFAULT_AI_URL = "http://localhost:8080/v1/chat/completions/";
  static final String DEFAULT_AI_APIKEY = "1234567";
  static final String DEFAULT_AI_IMG_MODEL = "stablediffusion";
  static final String DEFAULT_AI_IMG_URL = "http://localhost:8080/v1/images/generations/";
  static final String DEFAULT_AI_IMG_APIKEY = "1234567";
  static final String DEFAULT_AI_TTS_MODEL = "voice-de-eva_k-x-low";
  static final String DEFAULT_AI_TTS_URL = "http://localhost:8080/tts/";
  static final String DEFAULT_AI_TTS_APIKEY = "1234567";

  public static final Color GRAMMAR_COLOR_LT = new Color(255, 100, 0);
  public static final Color STYLE_COLOR_WT = new Color(0, 100, 0);
  public static final Color HINT_COLOR_WT = new Color(150, 150, 0);
  public static final Color STYLE_COLOR_BLUE = new Color(70, 80, 255);
  public static final Color HINT_COLOR_BLUE = new Color(150, 160, 255);
  public static final Color GRAMMAR_COLOR_DARK = new Color(100, 150, 255);
  public static final Color STYLE_COLOR_DARK = new Color(0, 140, 0);
  public static final Color HINT_COLOR_DARK = new Color(100, 100, 0);

  private static final String CONFIG_FILE = ".languagetool.cfg";

  private static final String CURRENT_PROFILE_KEY = "currentProfile";
  private static final String DEFINED_PROFILES_KEY = "definedProfiles";
  
  private static final String DISABLED_RULES_KEY = "disabledRules";
  private static final String ENABLED_RULES_KEY = "enabledRules";
  private static final String DISABLED_CATEGORIES_KEY = "disabledCategories";
  private static final String ENABLED_CATEGORIES_KEY = "enabledCategories";
  private static final String ENABLED_RULES_ONLY_KEY = "enabledRulesOnly";
  private static final String LANGUAGE_KEY = "language";
  private static final String MOTHER_TONGUE_KEY = "motherTongue";
  private static final String FIXED_LANGUAGE_KEY = "fixedLanguage";
  private static final String NGRAM_DIR_KEY = "ngramDir";
  private static final String AUTO_DETECT_KEY = "autoDetect";
  private static final String TAGGER_SHOWS_DISAMBIG_LOG_KEY = "taggerShowsDisambigLog";
  private static final String SERVER_RUN_KEY = "serverMode";
  private static final String SERVER_PORT_KEY = "serverPort";
  private static final String NO_DEFAULT_CHECK_KEY = "noDefaultCheck";
  private static final String PARA_CHECK_KEY = "numberParagraphs";
  private static final String RESET_CHECK_KEY = "doResetCheck";
  private static final String USE_QUEUE_KEY = "useTextLevelQueue";
  private static final String NO_BACKGROUND_CHECK_KEY = "noBackgroundCheck";
  private static final String USE_DOC_LANG_KEY = "useDocumentLanguage";
  private static final String USE_GUI_KEY = "useGUIConfig";
  private static final String FONT_NAME_KEY = "font.name";
  private static final String FONT_STYLE_KEY = "font.style";
  private static final String FONT_SIZE_KEY = "font.size";
  private static final String LF_NAME_KEY = "lookAndFeelName";
  private static final String COLOR_SELECTION_KEY = "colorSelection";
  private static final String CHECK_DIRECT_SPEACH_KEY = "checkDirectSpeech";
  private static final String THEME_SELECTION_KEY = "themeSelection";
  private static final String ERROR_COLORS_KEY = "errorColors";
  private static final String UNDERLINE_DEFAULT_COLORS_KEY = "underlineDefaultColors";
  private static final String UNDERLINE_COLORS_KEY = "underlineColors";
  private static final String UNDERLINE_RULE_COLORS_KEY = "underlineRuleColors";
  private static final String UNDERLINE_TYPES_KEY = "underlineTypes";
  private static final String UNDERLINE_RULE_TYPES_KEY = "underlineRuleTypes";
  private static final String CONFIGURABLE_RULE_VALUES_KEY = "configurableRuleValues";
  private static final String LT_SWITCHED_OFF_KEY = "ltSwitchedOff";
  private static final String IS_MULTI_THREAD_LO_KEY = "isMultiThread";
  private static final String EXTERNAL_RULE_DIRECTORY = "extRulesDirectory";
  private static final String DO_REMOTE_CHECK_KEY = "doRemoteCheck";
  private static final String OTHER_SERVER_URL_KEY = "otherServerUrl";
  private static final String REMOTE_USERNAME_KEY = "remoteUserName";
  private static final String REMOTE_APIKEY_KEY = "remoteApiKey";
  private static final String USE_OTHER_SERVER_KEY = "useOtherServer";
  private static final String IS_PREMIUM_KEY = "isPremium";
  private static final String MARK_SINGLE_CHAR_BOLD_KEY = "markSingleCharBold";
  private static final String LOG_LEVEL_KEY = "logLevel";
  private static final String USE_LT_SPELL_CHECKER_KEY = "UseLtSpellChecker";
  private static final String USE_LONG_MESSAGES_KEY = "UseLongMessages";
  private static final String NO_SYNONYMS_AS_SUGGESTIONS_KEY = "noSynonymsAsSuggestions";
  private static final String INCLUDE_TRACKED_CHANGES_KEY = "includeTrackedChanges";
  private static final String ENABLE_TMP_OFF_RULES_KEY = "enableTmpOffRules";
  private static final String ENABLE_GOAL_SPECIFIC_RULES_KEY = "enableGoalSpecificRules";
  private static final String FILTER_OVERLAPPING_MATCHES_KEY = "filterOverlappingMatches";
  private static final String SAVE_LO_CACHE_KEY = "saveLoCache";
  private static final String LT_VERSION_KEY = "ltVersion";
  private static final String AI_URL_KEY = "aiUrl";
  private static final String AI_APIKEY_KEY = "aiApiKey";
  private static final String AI_MODEL_KEY = "aiModel";
  private static final String AI_USE_AI_SUPPORT_KEY = "useAiSupport";
  private static final String AI_AUTO_CORRECT_KEY = "aiAutoCorrect";
  private static final String AI_AUTO_SUGGESTION_KEY = "aiAutoSuggestion";
  private static final String AI_SHOW_STYLISTIC_CHANGES_KEY = "aiShowStylisticChangesInt";
  private static final String AI_IMG_URL_KEY = "aiImgUrl";
  private static final String AI_IMG_APIKEY_KEY = "aiImgApiKey";
  private static final String AI_IMG_MODEL_KEY = "aiImgModel";
  private static final String AI_USE_AI_IMG_SUPPORT_KEY = "useAiImgSupport";
  private static final String AI_TTS_URL_KEY = "aiTtsUrl";
  private static final String AI_TTS_APIKEY_KEY = "aiTtsApiKey";
  private static final String AI_TTS_MODEL_KEY = "aiTtsModel";
  private static final String AI_USE_AI_TTS_SUPPORT_KEY = "useAiTtsSupport";

  private static final String DELIMITER = ",";
  // find all comma followed by zero or more white space characters that are preceded by ":" AND a valid 6-digit hex code
  // example: ":#44ffee,"
  private static final String COLOR_SPLITTER_REGEXP = "(?<=:#[0-9A-Fa-f]{6}),\\s*";
  //find all colon followed by a valid 6-digit hex code, e.g., ":#44ffee"
  private static final String COLOR_SPLITTER_REGEXP_COLON = ":(?=#[0-9A-Fa-f]{6})";
  // find all comma followed by zero or more white space characters that are preceded by at least one digit
  // example: "4,"
  private static final String CONFIGURABLE_RULE_SPLITTER_REGEXP = "(?<=[0-9]),\\s*";

  private static final String BLANK = "[ \t]";
  private static final String BLANK_REPLACE = "_";
  private static final String PROFILE_DELIMITER = "__";
  private static final String COMMA_REPLACE = "__comma__";
  
  // For new Maps, Sets or Lists add a clear to initOptions
  private final Map<String, String> configForOtherProfiles = new HashMap<>();
  private final Map<String, String> configForOtherLanguages = new HashMap<>();
  private final Map<ITSIssueType, Color> errorColors = new EnumMap<>(ITSIssueType.class);
  private final Map<String, Color> underlineColors = new HashMap<>();
  private final Map<String, Color> underlineRuleColors = new HashMap<>();
  private final Map<String, Short> underlineTypes = new HashMap<>();
  private final Map<String, Short> underlineRuleTypes = new HashMap<>();
  private final Map<String, Object[]> configurableRuleValues = new HashMap<>();
  private final Set<String> styleLikeCategories = new HashSet<>();
  private final Set<String> optionalCategories = new HashSet<>();
  private final Set<String> optionalRules = new HashSet<>();
  private final Map<String, String> specialTabCategories = new HashMap<>();

  // For new Maps, Sets or Lists add a clear to initOptions
  private final Set<String> disabledRuleIds = new HashSet<>();
  private final Set<String> enabledRuleIds = new HashSet<>();
  private final Set<String> disabledCategoryNames = new HashSet<>();
  private final Set<String> enabledCategoryNames = new HashSet<>();
  private final List<String> definedProfiles = new ArrayList<>();
  private final List<String> allProfileKeys = new ArrayList<>();
  private final List<String> allProfileLangKeys = new ArrayList<>();
  private final List<Color> underlineDefaultColors = new ArrayList<>();

  // Add new option default parameters to initOptions
  private Language lang;
  private File configFile;
  private boolean enabledRulesOnly = false;
  private Language language;
  private Language motherTongue = null;
  private Language fixedLanguage = null;
  private File ngramDirectory;
  private boolean runServer;
  private boolean autoDetect;
  private boolean taggerShowsDisambigLog;
  private boolean guiConfig;
  private String fontName;
  private int fontStyle = FONT_STYLE_INVALID;
  private int fontSize = FONT_SIZE_INVALID;
  private int serverPort = DEFAULT_SERVER_PORT;
  private int numParasToCheck = DEFAULT_NUM_CHECK_PARAS;
  private int colorSelection = DEFAULT_COLOR_SELECTION;
  private int checkDirectSpeech = DEFAULT_CHECK_DIRECT_SPEECH;
  private int themeSelection = DEFAULT_THEME_SELECTION;
  private boolean doResetCheck = DEFAULT_DO_RESET;
  private boolean isMultiThreadLO = DEFAULT_MULTI_THREAD;
  private boolean noBackgroundCheck = DEFAULT_NO_BACKGROUND_CHECK;
  private boolean useTextLevelQueue = DEFAULT_USE_QUEUE;
  private boolean useDocLanguage = DEFAULT_USE_DOC_LANGUAGE;
  private boolean doRemoteCheck = DEFAULT_DO_REMOTE_CHECK;
  private boolean useOtherServer = DEFAULT_USE_OTHER_SERVER;
  private boolean isPremium = DEFAULT_IS_PREMIUM;
  private boolean markSingleCharBold = DEFAULT_MARK_SINGLE_CHAR_BOLD;
  private boolean useLtSpellChecker = DEFAULT_USE_LT_SPELL_CHECKER;
  private boolean useLongMessages = DEFAULT_USE_LONG_MESSAGES;
  private boolean noSynonymsAsSuggestions = DEFAULT_NO_SYNONYMS_AS_SUGGESTIONS;
  private boolean includeTrackedChanges = DEFAULT_INCLUDE_TRACKED_CHANGES;
  private boolean enableTmpOffRules = DEFAULT_ENABLE_TMP_OFF_RULES;
  private boolean enableGoalSpecificRules = DEFAULT_ENABLE_GOAL_SPECIFIC_RULES;
  private boolean filterOverlappingMatches = DEFAULT_FILTER_OVERLAPPING_MATCHES;
  private boolean saveLoCache = DEFAULT_SAVE_LO_CACHE;
  private String externalRuleDirectory;
  private String lookAndFeelName;
  private String currentProfile = null;
  private String otherServerUrl = null;
  private String remoteUsername = null;
  private String remoteApiKey = null;
  private String logLevel = null;
  private String ltVersion = null;
  private boolean switchOff = false;
  private boolean isOffice = false;
  private boolean isOpenOffice = false;
  private String aiUrl = DEFAULT_AI_URL;
  private String aiApiKey = DEFAULT_AI_APIKEY;
  private String aiModel = DEFAULT_AI_MODEL;
  private boolean useAiSupport = DEFAULT_USE_AI_SUPPORT;
  private boolean aiAutoCorrect = DEFAULT_AI_AUTO_CORRECT;
  private boolean aiAutoSuggestion = DEFAULT_AI_AUTO_SUGGESTION;
  private int aiShowStylisticChanges = DEFAULT_AI_SHOW_STYLISTIC_CHANGES;
  private String aiImgUrl = DEFAULT_AI_IMG_URL;
  private String aiImgApiKey = DEFAULT_AI_IMG_APIKEY;
  private String aiImgModel = DEFAULT_AI_IMG_MODEL;
  private boolean useAiImgSupport = DEFAULT_USE_AI_IMG_SUPPORT;
  private String aiTtsUrl = DEFAULT_AI_TTS_URL;
  private String aiTtsApiKey = DEFAULT_AI_TTS_APIKEY;
  private String aiTtsModel = DEFAULT_AI_TTS_MODEL;
  private boolean useAiTtsSupport = DEFAULT_USE_AI_TTS_SUPPORT;
  
  /**
   * Uses the configuration file from the default location.
   *
   * @param lang The language for the configuration, used to distinguish
   *             rules that are enabled or disabled per language.
   */
  public WtConfiguration(Language lang) throws IOException {
    this(new File(System.getProperty("user.home")), CONFIG_FILE, lang);
  }

  public WtConfiguration(File baseDir, Language lang) throws IOException {
    this(baseDir, CONFIG_FILE, lang);
  }

  public WtConfiguration(File baseDir, String filename, Language lang) throws IOException {
    this(baseDir, filename, lang, false);
  }

  public WtConfiguration(File baseDir, String filename, Language lang, boolean isOffice) throws IOException {
    // already fails silently if file doesn't exist in loadConfiguration, don't fail here either
    // can cause problem when starting LanguageTool server as a user without a home directory because of default arguments
    //if (baseDir == null || !baseDir.isDirectory()) {
    //  throw new IllegalArgumentException("Cannot open file " + filename + " in directory " + baseDir);
    //}
    initOptions();
    this.lang = lang;
    this.isOffice = isOffice;
    this.isOpenOffice = isOffice && filename.contains("ooo");
    configFile = new File(baseDir, filename);
    setAllProfileKeys();
    loadConfiguration();
  }

  private WtConfiguration() {
    lang = null;
  }

  /**
   * Initialize variables and clears Maps, Sets and Lists
   */
  public void initOptions() {
    configForOtherLanguages.clear();
    underlineColors.clear();
    underlineRuleColors.clear();
    underlineDefaultColors.clear();
    underlineTypes.clear();
    underlineRuleTypes.clear();
    configurableRuleValues.clear();

    disabledRuleIds.clear();
    enabledRuleIds.clear();
    disabledCategoryNames.clear();
    enabledCategoryNames.clear();
    definedProfiles.clear();

    enabledRulesOnly = false;
    ngramDirectory = null;
    runServer = false;
    autoDetect = false;
    taggerShowsDisambigLog = false;
    guiConfig = false;
    fontName = null;
    fontStyle = FONT_STYLE_INVALID;
    fontSize = FONT_SIZE_INVALID;
    serverPort = DEFAULT_SERVER_PORT;
    numParasToCheck = DEFAULT_NUM_CHECK_PARAS;
    colorSelection = DEFAULT_COLOR_SELECTION;
    checkDirectSpeech = DEFAULT_CHECK_DIRECT_SPEECH;
    themeSelection = DEFAULT_THEME_SELECTION;
    doResetCheck = DEFAULT_DO_RESET;
    isMultiThreadLO = DEFAULT_MULTI_THREAD;
    noBackgroundCheck = DEFAULT_NO_BACKGROUND_CHECK;
    useTextLevelQueue = DEFAULT_USE_QUEUE;
    useDocLanguage = DEFAULT_USE_DOC_LANGUAGE;
    doRemoteCheck = DEFAULT_DO_REMOTE_CHECK;
    useOtherServer = DEFAULT_USE_OTHER_SERVER;
    isPremium = DEFAULT_IS_PREMIUM;
    markSingleCharBold = DEFAULT_MARK_SINGLE_CHAR_BOLD;
    useLtSpellChecker = DEFAULT_USE_LT_SPELL_CHECKER;
    useLongMessages = DEFAULT_USE_LONG_MESSAGES;
    noSynonymsAsSuggestions = DEFAULT_NO_SYNONYMS_AS_SUGGESTIONS;
    includeTrackedChanges = DEFAULT_INCLUDE_TRACKED_CHANGES;
    enableTmpOffRules = DEFAULT_ENABLE_TMP_OFF_RULES;
    enableGoalSpecificRules = DEFAULT_ENABLE_GOAL_SPECIFIC_RULES;
    filterOverlappingMatches = DEFAULT_FILTER_OVERLAPPING_MATCHES;
    saveLoCache = DEFAULT_SAVE_LO_CACHE;
    aiUrl = DEFAULT_AI_URL;
    aiApiKey = DEFAULT_AI_APIKEY;
    aiModel = DEFAULT_AI_MODEL;
    useAiSupport = DEFAULT_USE_AI_SUPPORT;
    aiAutoCorrect = DEFAULT_AI_AUTO_CORRECT;
    aiAutoSuggestion = DEFAULT_AI_AUTO_SUGGESTION;
    aiShowStylisticChanges = DEFAULT_AI_SHOW_STYLISTIC_CHANGES;
    aiImgUrl = DEFAULT_AI_IMG_URL;
    aiImgApiKey = DEFAULT_AI_IMG_APIKEY;
    aiImgModel = DEFAULT_AI_IMG_MODEL;
    useAiImgSupport = DEFAULT_USE_AI_IMG_SUPPORT;
    aiTtsUrl = DEFAULT_AI_TTS_URL;
    aiTtsApiKey = DEFAULT_AI_TTS_APIKEY;
    aiTtsModel = DEFAULT_AI_TTS_MODEL;
    useAiTtsSupport = DEFAULT_USE_AI_TTS_SUPPORT;
    externalRuleDirectory = null;
    lookAndFeelName = null;
    currentProfile = null;
    otherServerUrl = null;
    remoteUsername = null;
    remoteApiKey = null;
    logLevel = null;
    switchOff = false;
  }
  /**
   * Returns a copy of the given configuration.
   * @param configuration the object to copy.
   * @since LT2.6
   */
  public WtConfiguration copy(WtConfiguration configuration) {
    WtConfiguration copy = new WtConfiguration();
    copy.restoreState(configuration);
    return copy;
  }

  /**
   * Restore the state of this object from configuration.
   * @param configuration the object from which we will read the state
   * @since LT2.6
   */
  public void restoreState(WtConfiguration configuration) {
    this.configFile = configuration.configFile;
    this.language = configuration.language;
    this.lang = configuration.lang;
    this.motherTongue = configuration.motherTongue;
    this.fixedLanguage = configuration.fixedLanguage;
    this.ngramDirectory = configuration.ngramDirectory;
    this.runServer = configuration.runServer;
    this.autoDetect = configuration.autoDetect;
    this.taggerShowsDisambigLog = configuration.taggerShowsDisambigLog;
    this.guiConfig = configuration.guiConfig;
    this.fontName = configuration.fontName;
    this.fontStyle = configuration.fontStyle;
    this.fontSize = configuration.fontSize;
    this.serverPort = configuration.serverPort;
    this.numParasToCheck = configuration.numParasToCheck;
    this.colorSelection = configuration.colorSelection;
    this.checkDirectSpeech = configuration.checkDirectSpeech;
    this.themeSelection = configuration.themeSelection;
    this.doResetCheck = configuration.doResetCheck;
    this.useTextLevelQueue = configuration.useTextLevelQueue;
    this.noBackgroundCheck = configuration.noBackgroundCheck;
    this.isMultiThreadLO = configuration.isMultiThreadLO;
    this.useDocLanguage = configuration.useDocLanguage;
    this.lookAndFeelName = configuration.lookAndFeelName;
    this.externalRuleDirectory = configuration.externalRuleDirectory;
    this.currentProfile = configuration.currentProfile;
    this.doRemoteCheck = configuration.doRemoteCheck;
    this.useOtherServer = configuration.useOtherServer;
    this.isPremium = configuration.isPremium;
    this.markSingleCharBold = configuration.markSingleCharBold;
    this.useLtSpellChecker = configuration.useLtSpellChecker;
    this.useLongMessages = configuration.useLongMessages;
    this.noSynonymsAsSuggestions = configuration.noSynonymsAsSuggestions;
    this.includeTrackedChanges = configuration.includeTrackedChanges;
    this.enableTmpOffRules = configuration.enableTmpOffRules;
    this.enableGoalSpecificRules = configuration.enableGoalSpecificRules;
    this.filterOverlappingMatches = configuration.filterOverlappingMatches;
    this.saveLoCache = configuration.saveLoCache;
    this.otherServerUrl = configuration.otherServerUrl;
    this.remoteUsername = configuration.remoteUsername;
    this.remoteApiKey = configuration.remoteApiKey;
    this.logLevel = configuration.logLevel;
    this.isOffice = configuration.isOffice;
    this.isOpenOffice = configuration.isOpenOffice;
    this.ltVersion = configuration.ltVersion;
    this.aiUrl = configuration.aiUrl;
    this.aiApiKey = configuration.aiApiKey;
    this.aiModel = configuration.aiModel;
    this.useAiSupport = configuration.useAiSupport;
    this.aiAutoCorrect = configuration.aiAutoCorrect;
    this.aiAutoSuggestion = configuration.aiAutoSuggestion;
    this.aiShowStylisticChanges = configuration.aiShowStylisticChanges;
    this.aiImgUrl = configuration.aiImgUrl;
    this.aiImgApiKey = configuration.aiImgApiKey;
    this.aiImgModel = configuration.aiImgModel;
    this.useAiImgSupport = configuration.useAiImgSupport;
    this.aiTtsUrl = configuration.aiTtsUrl;
    this.aiTtsApiKey = configuration.aiTtsApiKey;
    this.aiTtsModel = configuration.aiTtsModel;
    this.useAiTtsSupport = configuration.useAiTtsSupport;
    
    this.disabledRuleIds.clear();
    this.disabledRuleIds.addAll(configuration.disabledRuleIds);
    this.enabledRuleIds.clear();
    this.enabledRuleIds.addAll(configuration.enabledRuleIds);
    this.disabledCategoryNames.clear();
    this.disabledCategoryNames.addAll(configuration.disabledCategoryNames);
    this.enabledCategoryNames.clear();
    this.enabledCategoryNames.addAll(configuration.enabledCategoryNames);
    this.underlineDefaultColors.clear();
    this.underlineDefaultColors.addAll(configuration.underlineDefaultColors);
    this.configForOtherLanguages.clear();
    for (String key : configuration.configForOtherLanguages.keySet()) {
      this.configForOtherLanguages.put(key, configuration.configForOtherLanguages.get(key));
    }
    this.underlineColors.clear();
    for (Map.Entry<String, Color> entry : configuration.underlineColors.entrySet()) {
      this.underlineColors.put(entry.getKey(), entry.getValue());
    }
    this.underlineRuleColors.clear();
    for (Map.Entry<String, Color> entry : configuration.underlineRuleColors.entrySet()) {
      this.underlineRuleColors.put(entry.getKey(), entry.getValue());
    }
    this.underlineTypes.clear();
    for (Map.Entry<String, Short> entry : configuration.underlineTypes.entrySet()) {
      this.underlineTypes.put(entry.getKey(), entry.getValue());
    }
    this.underlineRuleTypes.clear();
    for (Map.Entry<String, Short> entry : configuration.underlineRuleTypes.entrySet()) {
      this.underlineRuleTypes.put(entry.getKey(), entry.getValue());
    }
    this.configurableRuleValues.clear();
    for (Map.Entry<String, Object[]> entry : configuration.configurableRuleValues.entrySet()) {
      this.configurableRuleValues.put(entry.getKey(), entry.getValue());
    }
    this.styleLikeCategories.clear();
    this.styleLikeCategories.addAll(configuration.styleLikeCategories);
    this.optionalCategories.clear();
    this.optionalCategories.addAll(configuration.optionalCategories);
    this.optionalRules.clear();
    this.optionalRules.addAll(configuration.optionalRules);
    this.specialTabCategories.clear();
    for (Map.Entry<String, String> entry : configuration.specialTabCategories.entrySet()) {
      this.specialTabCategories.put(entry.getKey(), entry.getValue());
    }
    this.definedProfiles.clear();
    this.definedProfiles.addAll(configuration.definedProfiles);
    this.allProfileLangKeys.clear();
    this.allProfileLangKeys.addAll(configuration.allProfileLangKeys);
    this.allProfileKeys.clear();
    this.allProfileKeys.addAll(configuration.allProfileKeys);
    this.configForOtherProfiles.clear();
    for (Entry<String, String> entry : configuration.configForOtherProfiles.entrySet()) {
      this.configForOtherProfiles.put(entry.getKey(), entry.getValue());
    }
  }

  public void setConfigFile(File configFile) {
    this.configFile = configFile;
  }

  public Set<String> getDisabledRuleIds() {
    return disabledRuleIds;
  }

  public Set<String> getEnabledRuleIds() {
    return enabledRuleIds;
  }

  public Set<String> getDisabledCategoryNames() {
    return disabledCategoryNames;
  }

  public Set<String> getEnabledCategoryNames() {
    return enabledCategoryNames;
  }

  public void setDisabledRuleIds(Set<String> ruleIds) {
    disabledRuleIds.clear();
    disabledRuleIds.addAll(ruleIds);
    enabledRuleIds.removeAll(ruleIds);
  }

  public void addDisabledRuleIds(Set<String> ruleIds) {
    disabledRuleIds.addAll(ruleIds);
    enabledRuleIds.removeAll(ruleIds);
  }

  public void removeDisabledRuleIds(Set<String> ruleIds) {
    disabledRuleIds.removeAll(ruleIds);
    enabledRuleIds.addAll(ruleIds);
  }

  public void removeDisabledRuleId(String ruleId) {
    disabledRuleIds.remove(ruleId);
  }

  public void removeEnabledRuleId(String ruleId) {
    enabledRuleIds.remove(ruleId);
  }

  public void setEnabledRuleIds(Set<String> ruleIds) {
    enabledRuleIds.clear();
    enabledRuleIds.addAll(ruleIds);
  }

  public void setDisabledCategoryNames(Set<String> categoryNames) {
    disabledCategoryNames.clear();
    disabledCategoryNames.addAll(categoryNames);
  }

  public void setEnabledCategoryNames(Set<String> categoryNames) {
    enabledCategoryNames.clear();
    enabledCategoryNames.addAll(categoryNames);
  }

  public boolean getEnabledRulesOnly() {
    return enabledRulesOnly;
  }

  public Language getLanguage() {
    return language;
  }

  public void setLanguage(Language language) {
    this.language = language;
  }

  public Language getMotherTongue() {
    return motherTongue;
  }

  public void setMotherTongue(Language motherTongue) {
    this.motherTongue = motherTongue;
  }

  public Language getFixedLanguage() {
    return fixedLanguage;
  }

  public void setFixedLanguage(Language fixedLanguage) {
    this.fixedLanguage = fixedLanguage;
  }

  public Language getDefaultLanguage() {
    if (useDocLanguage) {
      return null;
    }
    return fixedLanguage;
  }

  public void setUseDocLanguage(boolean useDocLang) {
    useDocLanguage = useDocLang;
  }

  public boolean getUseDocLanguage() {
    return useDocLanguage;
  }

  public boolean getAutoDetect() {
    return autoDetect;
  }

  public void setAutoDetect(boolean autoDetect) {
    this.autoDetect = autoDetect;
  }

  public void setRemoteCheck(boolean doRemoteCheck) {
    this.doRemoteCheck = doRemoteCheck;
  }

  public boolean doRemoteCheck() {
    return doRemoteCheck;
  }

  public void setUseOtherServer(boolean useOtherServer) {
    this.useOtherServer = useOtherServer;
  }

  public boolean useOtherServer() {
    return useOtherServer;
  }

  public void setPremium(boolean isPremium) {
    this.isPremium = isPremium;
  }

  public boolean isPremium() {
    return isPremium;
  }

  public void setOtherServerUrl(String otherServerUrl) {
    this.otherServerUrl = otherServerUrl;
  }

  public String getServerUrl() {
    return useOtherServer ? otherServerUrl : null;
  }

  public void setRemoteUsername(String remoteUsername) {
    this.remoteUsername = remoteUsername;
  }

  public String getRemoteUsername() {
    return isPremium ? remoteUsername : null;
  }

  public void setRemoteApiKey(String remoteApiKey) {
    this.remoteApiKey = remoteApiKey;
  }

  public String aiUrl() {
    return aiUrl;
  }

  public void setAiUrl(String aiUrl) {
    this.aiUrl = aiUrl;
  }

  public String aiModel() {
    return aiModel;
  }

  public void setAiModel(String aiModel) {
    this.aiModel = aiModel;
  }

  public String aiApiKey() {
    return aiApiKey;
  }

  public void setAiApiKey(String aiApiKey) {
    this.aiApiKey = aiApiKey;
  }

  public boolean useAiSupport() {
    return useAiSupport;
  }

  public void setUseAiSupport(boolean useAiSupport) {
    this.useAiSupport = useAiSupport;
  }

  public boolean aiAutoCorrect() {
    return aiAutoCorrect;
  }

  public void setAiAutoCorrect(boolean aiAutoCorrect) {
    this.aiAutoCorrect = aiAutoCorrect;
  }

  public boolean aiAutoSuggestion() {
    return aiAutoSuggestion;
  }

  public void setAiAutoSuggestion(boolean aiAutoSuggestion) {
    this.aiAutoSuggestion = aiAutoSuggestion;
  }

  public int aiShowStylisticChanges() {
    return aiShowStylisticChanges;
  }

  public void setAiShowStylisticChanges(int aiShowStylisticChanges) {
    this.aiShowStylisticChanges = aiShowStylisticChanges;
  }

  public String aiImgUrl() {
    return aiImgUrl;
  }

  public void setAiImgUrl(String aiImgUrl) {
    this.aiImgUrl = aiImgUrl;
  }

  public String aiImgModel() {
    return aiImgModel;
  }

  public void setAiImgModel(String aiImgModel) {
    this.aiImgModel = aiImgModel;
  }

  public String aiImgApiKey() {
    return aiImgApiKey;
  }

  public void setAiImgApiKey(String aiImgApiKey) {
    this.aiImgApiKey = aiImgApiKey;
  }

  public boolean useAiImgSupport() {
    return useAiImgSupport;
  }

  public void setUseAiImgSupport(boolean useAiImgSupport) {
    this.useAiImgSupport = useAiImgSupport;
  }

  public String aiTtsUrl() {
    return aiTtsUrl;
  }

  public void setAiTtsUrl(String aiTtsUrl) {
    this.aiTtsUrl = aiTtsUrl;
  }

  public String aiTtsModel() {
    return aiTtsModel;
  }

  public void setAiTtsModel(String aiTtsModel) {
    this.aiTtsModel = aiTtsModel;
  }

  public String aiTtsApiKey() {
    return aiTtsApiKey;
  }

  public void setAiTtsApiKey(String aiTtsApiKey) {
    this.aiTtsApiKey = aiTtsApiKey;
  }

  public boolean useAiTtsSupport() {
    return useAiTtsSupport;
  }

  public void setUseAiTtsSupport(boolean useAiTtsSupport) {
    this.useAiTtsSupport = useAiTtsSupport;
  }

  public String getRemoteApiKey() {
    return isPremium ? remoteApiKey : null;
  }

  public String getlogLevel() {
    return logLevel;
  }

  public void setMarkSingleCharBold(boolean markSingleCharBold) {
    this.markSingleCharBold = markSingleCharBold;
  }

  public boolean markSingleCharBold() {
    return markSingleCharBold;
  }
  
  public void setUseLtSpellChecker(boolean useLtSpellChecker) {
    this.useLtSpellChecker = useLtSpellChecker;
  }

  public boolean useLtSpellChecker() {
    return useLtSpellChecker;
  }
  
  public void setUseLongMessages(boolean useLongMessages) {
    this.useLongMessages = useLongMessages;
  }

  public boolean useLongMessages() {
    return useLongMessages;
  }
  
  public void setNoSynonymsAsSuggestions(boolean noSynonymsAsSuggestions) {
    this.noSynonymsAsSuggestions = noSynonymsAsSuggestions;
  }

  public boolean noSynonymsAsSuggestions() {
    return noSynonymsAsSuggestions;
  }
  
  public void setIncludeTrackedChanges(boolean includeTrackedChanges) {
    this.includeTrackedChanges = includeTrackedChanges;
  }

  public boolean includeTrackedChanges() {
    return includeTrackedChanges;
  }
  
  public void setEnableTmpOffRules(boolean enableTmpOffRules) {
    this.enableTmpOffRules = enableTmpOffRules;
  }

  public boolean enableTmpOffRules() {
    return enableTmpOffRules;
  }
  
  public void setEnableGoalSpecificRules(boolean enableGoalSpecificRules) {
    this.enableGoalSpecificRules = enableGoalSpecificRules;
  }

  public boolean enableGoalSpecificRules() {
    return enableGoalSpecificRules;
  }
  
  public void setFilterOverlappingMatches(boolean filterOverlappingMatches) {
    this.filterOverlappingMatches = filterOverlappingMatches;
  }

  public boolean filterOverlappingMatches() {
    return filterOverlappingMatches;
  }
  
  public void setSaveLoCache(boolean saveLoCache) {
    this.saveLoCache = saveLoCache;
  }

  public boolean saveLoCache() {
    return saveLoCache;
  }
  
  /**
   * Determines whether the tagger window will also print the disambiguation
   * log.
   * @return true if the tagger window will print the disambiguation log,
   * false otherwise
   * @since LT3.3
   */
  public boolean getTaggerShowsDisambigLog() {
    return taggerShowsDisambigLog;
  }

  /**
   * Enables or disables the disambiguation log on the tagger window,
   * depending on the value of the parameter taggerShowsDisambigLog.
   * @param taggerShowsDisambigLog If true, the tagger window will print the
   * @since LT3.3
   */
  public void setTaggerShowsDisambigLog(boolean taggerShowsDisambigLog) {
    this.taggerShowsDisambigLog = taggerShowsDisambigLog;
  }

  public boolean getRunServer() {
    return runServer;
  }

  public void setRunServer(boolean runServer) {
    this.runServer = runServer;
  }

  public int getServerPort() {
    return serverPort;
  }

  public void setUseGUIConfig(boolean useGUIConfig) {
    this.guiConfig = useGUIConfig;
  }

  public boolean getUseGUIConfig() {
    return guiConfig;
  }

  public void setServerPort(int serverPort) {
    this.serverPort = serverPort;
  }

  public String getExternalRuleDirectory() {
    return externalRuleDirectory;
  }

  public void setExternalRuleDirectory(String path) {
    externalRuleDirectory = path;
  }

  /**
   * get the number of paragraphs to be checked for TextLevelRules
   * @since LT4.0
   */
  public int getNumParasToCheck() {
    return numParasToCheck;
  }

  /**
   * set the number of paragraphs to be checked for TextLevelRules
   * @since LT4.0
   */
  public void setNumParasToCheck(int numParas) {
    this.numParasToCheck = numParas;
  }

  /**
   * get the color model selected
   * @since WT 1.0
   */
  public int getColorSelection() {
    return colorSelection;
  }

  /**
   * set the color model to use
   * @since WT 1.0
   */
  public void setColorSelection(int colorSelection) {
    this.colorSelection = colorSelection;
  }

  /**
   * get the option to check the text inside direct speech or not
   * @since  1.2
   */
  public int getCheckDirectSpeech() {
    return checkDirectSpeech;
  }

  /**
   * get the option to check the text inside direct speech or not
   * @since  1.2
   */
  public void setCheckDirectSpeech(int checkDirectSpeech) {
    this.checkDirectSpeech = checkDirectSpeech;
  }

  /**
   * get the theme selected
   * @since 1.2
   */
  public int getThemeSelection() {
    return themeSelection;
  }

  /**
   * set the theme to use
   * @since 1.2
   */
  public void setThemeSelection(int themeSelection) {
    this.themeSelection = themeSelection;
  }

  /**
   * will all paragraphs check after every change of text?
   * @since LT4.2
   */
  public boolean isResetCheck() {
    return doResetCheck;
  }

  /**
   * set all paragraphs to be checked after every change of text
   * @since LT4.2
   */
  public void setDoResetCheck(boolean resetCheck) {
    this.doResetCheck = resetCheck;
  }

  /**
   * will all paragraphs not checked after every change of text 
   * if more than one document loaded?
   * @since LT4.5
   */
  public boolean useTextLevelQueue() {
    return useTextLevelQueue;
  }

  /**
   * set all paragraphs to be not checked after every change of text
   * if more than one document loaded?
   * @since LT4.5
   */
  public void setUseTextLevelQueue(boolean useTextLevelQueue) {
    this.useTextLevelQueue = useTextLevelQueue;
  }

  /**
   * set option to switch off background check
   * if true: LT engine is switched of (no marks inside of document)
   * @since LT5.2
   */
  public void setNoBackgroundCheck(boolean noBackgroundCheck) {
    this.noBackgroundCheck = noBackgroundCheck;
  }

  /**
   * set option to switch off background check
   * and save configuration
   * @since LT5.2
   */
  public void saveNoBackgroundCheck(boolean noBackgroundCheck, Language lang) throws IOException {
    this.noBackgroundCheck = noBackgroundCheck;
    saveConfiguration(lang);
  }
  
  /**
   * return true if background check is switched of
   * (no marks inside of document)
   * @since LT5.2
   */
  public boolean noBackgroundCheck() {
    return noBackgroundCheck;
  }
  
  /**
   * get the current profile
   * @since LT4.7
   */
  public String getCurrentProfile() {
    return currentProfile;
  }
  
  /**
   * set the current profile
   * @since LT4.7
   */
  public void setCurrentProfile(String profile) {
    currentProfile = profile;
  }
  
  /**
   * get the current profile
   * @since LT4.7
   */
  public List<String> getDefinedProfiles() {
    return definedProfiles;
  }
  
  /**
   * add a new profile
   * @since LT4.7
   */
  public void addProfile(String profile) {
    definedProfiles.add(profile);
  }
  
  /**
   * add a list of profiles
   * @since LT4.7
   */
  public void addProfiles(List<String> profiles) {
    definedProfiles.clear();
    definedProfiles.addAll(profiles);
  }
  
  /**
   * remove an existing profile
   * @since LT4.7
   */
  public void removeProfile(String profile) {
    definedProfiles.remove(profile);
  }

  /**
   * run LO in multi thread mode
   * @since LT4.6
   */
  public void setMultiThreadLO(boolean isMultiThread) {
    this.isMultiThreadLO = isMultiThread;
  }

  /**
   * shall LO run in multi thread mode 
   * @since LT4.6
   */
  public boolean isMultiThread() {
    return isMultiThreadLO;
  }

  /**
   * Returns the name of the GUI's editing textarea font.
   * @return the name of the font.
   * @see Font#getFamily()
   * @since LT2.6
   */
  public String getFontName() {
    return fontName;
  }

  /**
   * Sets the name of the GUI's editing textarea font.
   * @param fontName the name of the font.
   * @see Font#getFamily()
   * @since LT2.6
   */
  public void setFontName(String fontName) {
    this.fontName = fontName;
  }

  /**
   * Returns the style of the GUI's editing textarea font.
   * @return the style of the font.
   * @see Font#getStyle()
   * @since LT2.6
   */
  public int getFontStyle() {
    return fontStyle;
  }

  /**
   * Sets the style of the GUI's editing textarea font.
   * @param fontStyle the style of the font.
   * @see Font#getStyle()
   * @since LT2.6
   */
  public void setFontStyle(int fontStyle) {
    this.fontStyle = fontStyle;
  }

  /**
   * Returns the size of the GUI's editing textarea font.
   * @return the size of the font.
   * @see Font#getSize()
   * @since LT2.6
   */
  public int getFontSize() {
    return fontSize;
  }

  /**
   * Sets the size of the GUI's editing textarea font.
   * @param fontSize the size of the font.
   * @see Font#getSize()
   * @since LT2.6
   */
  public void setFontSize(int fontSize) {
    this.fontSize = fontSize;
  }

  /**
   * Returns the name of the GUI's LaF.
   * @return the name of the LaF.
   * @see javax.swing.UIManager.LookAndFeelInfo#getName()
   * @since LT2.6
   */
  public String getLookAndFeelName() {
    return this.lookAndFeelName;
  }

  /**
   * Sets the name of the GUI's LaF.
   * @param lookAndFeelName the name of the LaF.
   * @see javax.swing.UIManager.LookAndFeelInfo#getName()
   * @since LT2.6 @see
   */
  public void setLookAndFeelName(String lookAndFeelName) {
    this.lookAndFeelName = lookAndFeelName;
  }

  /**
   * Directory with ngram data or null.
   * @since LT3.0
   */
  @Nullable
  public File getNgramDirectory() {
    return ngramDirectory;
  }

  /**
   * Sets the directory with ngram data (may be null).
   * @since LT3.0
   */
  public void setNgramDirectory(File dir) {
    this.ngramDirectory = dir;
  }

  /**
   * @since LT2.8
   */
  public Map<ITSIssueType, Color> getErrorColors() {
    return errorColors;
  }

  /**
   * @since LT4.3
   * Returns true if category is style like
   */
  public boolean isStyleCategory(String category) {
    return styleLikeCategories.contains(category);
  }

  /**
   * @since 1.2
   * Returns true if category is style like
   */
  public boolean isOptionalCategory(String category) {
    return optionalCategories.contains(category);
  }

  /**
   * @since LT4.4
   * Initialize set of style like categories
   */
  public void initStyleCategories(List<Rule> allRules) {
    Map<String, Boolean> categoryMap = new HashMap<>();
    for (Rule rule : allRules) {
      if (rule.getCategory().getTabName() != null && !specialTabCategories.containsKey(rule.getCategory().getName())) {
        specialTabCategories.put(rule.getCategory().getName(), rule.getCategory().getTabName());
      }
      if (rule.getLocQualityIssueType().toString().equalsIgnoreCase("STYLE")
              || rule.getLocQualityIssueType().toString().equalsIgnoreCase("REGISTER")
//              || rule.getCategory().getId().toString().equals("TYPOGRAPHY")
              || rule.getCategory().getId().toString().equals("STYLE")) {
        styleLikeCategories.add(rule.getCategory().getName());
      }
      boolean isDefault = !rule.isOfficeDefaultOff() && (!rule.isDefaultOff() || rule.isOfficeDefaultOn());
      if (!isDefault) {
        optionalRules.add(rule.getId());
      }
      if (rule.getCategory().isDefaultOff()) {
        isDefault = false;
      }
      if (!categoryMap.containsKey(rule.getCategory().getName()) || isDefault) {
        categoryMap.put(rule.getCategory().getName(), isDefault);
      }
    }
    for (String category : categoryMap.keySet()) {
      if (!categoryMap.get(category)) {
        optionalCategories.add(category);
      }
    }
    optionalRules.add(WtOfficeTools.AI_GRAMMAR_OTHER_RULE_ID);
    optionalCategories.add(WtOfficeTools.AI_STYLE_CATEGORY);
  }

  /**
   * @since LT4.3
   * Returns true if category is a special Tab category
   */
  public boolean isSpecialTabCategory(String category) {
    return specialTabCategories.containsKey(category);
  }

  /**
   * @since LT4.3
   * Returns true if category is member of named special Tab
   */
  public boolean isInSpecialTab(String category, String tabName) {
    if (specialTabCategories.containsKey(category)) {
      return specialTabCategories.get(category).equals(tabName);
    }
    return false;
  }

  /**
   * @since LT4.3
   * Returns all special tab names
   */
  public String[] getSpecialTabNames() {
    Set<String> tabNames = new HashSet<>();
    for (Map.Entry<String, String> entry : specialTabCategories.entrySet()) {
      tabNames.add(entry.getValue());
    }
    return tabNames.toArray(new String[0]);
  }

  /**
   * @since LT4.3
   * Returns all categories for a named special tab
   */
  public Set<String> getSpecialTabCategories(String tabName) {
    Set<String> tabCategories = new HashSet<>();
    for (Map.Entry<String, String> entry : specialTabCategories.entrySet()) {
      if (entry.getKey().equals(tabName)) {
        tabCategories.add(entry.getKey());
      }
    }
    return tabCategories;
  }

  /**
   * @since LT4.2
   */
  public Map<String, Color> getUnderlineColors() {
    return underlineColors;
  }

  /**
   * @since LT5.3
   */
  public Map<String, Color> getUnderlineRuleColors() {
    return underlineRuleColors;
  }

  /**
   * @since LTWT 1.1
   */
  public List<Color> getUnderlineDefaultColors() {
    List<Color> colors = new ArrayList<>();
    if (underlineDefaultColors.size() != 3) {
      colors.add(Color.BLUE);
      colors.add(STYLE_COLOR_WT);
      colors.add(HINT_COLOR_WT);
    } else {
      colors.addAll(underlineDefaultColors);
    }
    return colors;
  }

  /**
   * @since LTWT 1.1
   * Set the rule default colors
   */
  public void setUnderlineDefaultColor(List<Color> colors) {
    underlineDefaultColors.clear();
    underlineDefaultColors.addAll(colors);
  }

  /**
   * @since LT4.2
   * Get the color to underline a rule match by the Name of its category
   */
  public Color getUnderlineColor(String category, String ruleId) {
    if (isOpenOffice) {
      return Color.blue;
    }
    if (ruleId != null && underlineRuleColors.containsKey(ruleId)) {
      return underlineRuleColors.get(ruleId);
    }
    if (underlineColors.containsKey(category)) {
      return underlineColors.get(category);
    }
//    boolean categoryIsDefault = !optionalCategories.contains(category);
    boolean categoryIsDefault = ruleId == null ? !optionalCategories.contains(category) : !optionalRules.contains(ruleId);
    boolean categoryIsStyle = styleLikeCategories.contains(category);
    if (colorSelection == 2) {
      if (!categoryIsDefault) {
        return STYLE_COLOR_BLUE;
      }
      if (styleLikeCategories.contains(category)) {
        return Color.blue;
      } else {
        return GRAMMAR_COLOR_LT;
      }
    } else {
      if (colorSelection == 99 && underlineDefaultColors.size() == 3) {
        return categoryIsDefault ? (!categoryIsStyle ? 
            underlineDefaultColors.get(0) : underlineDefaultColors.get(1)) : underlineDefaultColors.get(2);
      }
      if (!categoryIsDefault) {
        return colorSelection == 1 ? HINT_COLOR_BLUE : colorSelection == 3 ? HINT_COLOR_DARK : HINT_COLOR_WT;
      }
      if (categoryIsStyle) {
        return colorSelection == 1 ? STYLE_COLOR_BLUE : colorSelection == 3 ? STYLE_COLOR_DARK : STYLE_COLOR_WT;
      }
      return colorSelection == 3 ? GRAMMAR_COLOR_DARK : Color.blue;
    }
  }

  /**
   * @since LT4.2
   * Set the color to underline a rule match for its category
   */
  public void setUnderlineColor(String category, Color col) {
    underlineColors.put(category, col);
  }

  /**
   * @since LT5.3
   * Set the color to underline a rule match for this rule
   */
  public void setUnderlineRuleColor(String ruleId, Color col) {
    underlineRuleColors.put(ruleId, col);
  }

  /**
   * @since LT4.2
   * Set the category color back to default (removes category from map)
   */
  public void setDefaultUnderlineColor(String category) {
    underlineColors.remove(category);
  }

  /**
   * @since LT5.3
   * Set the rule color back to default (removes rule from map)
   */
  public void setDefaultUnderlineRuleColor(String ruleId) {
    underlineRuleColors.remove(ruleId);
  }

  /**
   * @since LT4.9
   */
  public Map<String, Short> getUnderlineTypes() {
    return underlineTypes;
  }

  /**
   * @since LT5.3
   */
  public Map<String, Short> getUnderlineRuleTypes() {
    return underlineRuleTypes;
  }

  /**
   * @since LT4.9
   * Get the type to underline a rule match by the Name of its category
   */
  public Short getUnderlineType(String category, String ruleId) {
    if (ruleId != null && underlineRuleTypes.containsKey(ruleId)) {
      return underlineRuleTypes.get(ruleId);
    }
    if (underlineTypes.containsKey(category)) {
      return underlineTypes.get(category);
    }
    return UNDERLINE_WAVE;
  }

  /**
   * @since LT4.9
   * Set the type to underline a rule match for its category
   */
  public void setUnderlineType(String category, short type) {
    underlineTypes.put(category, type);
  }

  /**
   * @since LT5.3
   * Set the type to underline a rule match
   */
  public void setUnderlineRuleType(String ruleID, short type) {
    underlineRuleTypes.put(ruleID, type);
  }

  /**
   * @since LT4.9
   * Set the type back to default (removes category from map)
   */
  public void setDefaultUnderlineType(String category) {
    underlineTypes.remove(category);
  }

  /**
   * @since LT5.3
   * Set the type back to default (removes ruleId from map)
   */
  public void setDefaultUnderlineRuleType(String ruleID) {
    underlineRuleTypes.remove(ruleID);
  }

  /**
   * returns all configured values
   * @since LT4.2
   */
  public Map<String, Object[]> getConfigurableValues() {
    return configurableRuleValues;
  }

  /**
   * Get the configurable value of a rule by ruleID
   * returns default value if no value is set by configuration
   * @since LT6.5
   */
  @SuppressWarnings("unchecked")
  public <T> T getConfigValueByID(String ruleID, int index, Class<T> clazz, T defaultValue) {
    Object[] value = configurableRuleValues.get(ruleID);
    if (value == null || index >= value.length || !clazz.isInstance(value[index])) {
      return defaultValue;
    }
    return (T) value[index];
  }

  /**
   * Set the values for a rule by ruleID
   * @since LT6.5
   */
  public void setConfigurableValue(String ruleID, Object[] values) {
    configurableRuleValues.put(ruleID, values);
  }

  /**
   * Remove the configuration values of a rule by ruleID
   * @since LT6.5
   */
  public void removeConfigurableValue(String ruleID) {
    configurableRuleValues.remove(ruleID);
  }

  /**
   * only single paragraph mode can be used (for OO and old LO
   * @since LT5.6
   */
  public boolean onlySingleParagraphMode() {
    return isOpenOffice;
  }

  /**
   * Set LT is switched Off or On
   * save configuration
   * @since LT4.4
   */
//  public void setSwitchedOff(boolean switchOff, Language lang) throws IOException {
//    this.switchOff = switchOff;
//    saveConfiguration(lang);
//  }
  
  /**
   * Test if http-server URL is correct
   */
  public boolean isValidServerUrl(String url) {
    if (url.endsWith("/") || url.endsWith("/v2") || !Pattern.matches("https?://.+(:\\d+)?.*", url)) {
      return false;
    }
    return true;
  }

  public boolean isValidAiServerUrl(String url) {
    if (!Pattern.matches("https?://.+(:\\d+)?.*", url)) {
      return false;
    }
    return true;
  }

  private void loadConfiguration() throws IOException {
    loadConfiguration(null);
  }

  public void loadConfiguration(String profile) throws IOException {
    String qualifier = getQualifier(lang);
    
    File cfgFile = configFile;

    try (FileInputStream fis = new FileInputStream(cfgFile)) {

      Properties props = new Properties();
      props.load(fis);
      
      if (profile == null) {
        String curProfileStr = (String) props.get(CURRENT_PROFILE_KEY);
        if (curProfileStr != null) {
          currentProfile = curProfileStr;
        }
      } else {
        currentProfile = profile;
      }
      definedProfiles.addAll(getListFromProperties(props, DEFINED_PROFILES_KEY));
      
      ltVersion = (String) props.get(LT_VERSION_KEY);
      
      if (ltVersion != null) {
        String motherTongueStr = (String) props.get(MOTHER_TONGUE_KEY);
        if (motherTongueStr != null && !motherTongueStr.equals("xx")) {
          motherTongue = Languages.getLanguageForShortCode(motherTongueStr);
        }
      }

      logLevel = (String) props.get(LOG_LEVEL_KEY);
      
      storeConfigForAllProfiles(props);
      
      String prefix;
      if (currentProfile == null) {
        prefix = "";
      } else {
        prefix = currentProfile;
      }
      if (!prefix.isEmpty()) {
        prefix = prefix.replaceAll(BLANK, BLANK_REPLACE);
        prefix += PROFILE_DELIMITER;
      }
      loadCurrentProfile(props, prefix, qualifier);
    } catch (FileNotFoundException e) {
      // file not found: okay, leave disabledRuleIds empty
    }
  }
  
  private void loadCurrentProfile(Properties props, String prefix, String qualifier) {
    
    String useDocLangString = (String) props.get(prefix + USE_DOC_LANG_KEY);
    if (useDocLangString != null) {
      useDocLanguage = Boolean.parseBoolean(useDocLangString);
    }
    if (ltVersion == null) {
      String motherTongueStr = (String) props.get(prefix + MOTHER_TONGUE_KEY);
      if (motherTongueStr != null && !motherTongueStr.equals("xx")) {
        if (isOffice) {
          fixedLanguage = Languages.getLanguageForShortCode(motherTongueStr);
        } else {
          motherTongue = Languages.getLanguageForShortCode(motherTongueStr);
        }
      }
    } else {
      String fixedLanguageStr = (String) props.get(prefix + FIXED_LANGUAGE_KEY);
      if (fixedLanguageStr != null) {
        fixedLanguage = Languages.getLanguageForShortCode(fixedLanguageStr);
      }
    }
    if (!useDocLanguage && fixedLanguage != null) {
      qualifier = getQualifier(fixedLanguage);
    }

    disabledRuleIds.addAll(getListFromProperties(props, prefix + DISABLED_RULES_KEY + qualifier));
    enabledRuleIds.addAll(getListFromProperties(props, prefix + ENABLED_RULES_KEY + qualifier));
    disabledCategoryNames.addAll(getListFromProperties(props, prefix + DISABLED_CATEGORIES_KEY + qualifier));
    enabledCategoryNames.addAll(getListFromProperties(props, prefix + ENABLED_CATEGORIES_KEY + qualifier));
    enabledRulesOnly = "true".equals(props.get(prefix + ENABLED_RULES_ONLY_KEY));

    String languageStr = (String) props.get(prefix + LANGUAGE_KEY);
    if (languageStr != null) {
      language = Languages.getLanguageForShortCode(languageStr);
    }
    String ngramDir = (String) props.get(prefix + NGRAM_DIR_KEY);
    if (ngramDir != null) {
      ngramDirectory = new File(ngramDir);
    }

    autoDetect = "true".equals(props.get(prefix + AUTO_DETECT_KEY));
    taggerShowsDisambigLog = "true".equals(props.get(prefix + TAGGER_SHOWS_DISAMBIG_LOG_KEY));
    guiConfig = "true".equals(props.get(prefix + USE_GUI_KEY));
    runServer = "true".equals(props.get(prefix + SERVER_RUN_KEY));

    fontName = (String) props.get(prefix + FONT_NAME_KEY);
    if (props.get(prefix + FONT_STYLE_KEY) != null) {
      try {
        fontStyle = Integer.parseInt((String) props.get(prefix + FONT_STYLE_KEY));
      } catch (NumberFormatException e) {
        // Ignore
      }
    }
    if (props.get(prefix + FONT_SIZE_KEY) != null) {
      try {
        fontSize = Integer.parseInt((String) props.get(prefix + FONT_SIZE_KEY));
      } catch (NumberFormatException e) {
        // Ignore
      }
    }
    lookAndFeelName = (String) props.get(prefix + LF_NAME_KEY);

    String serverPortString = (String) props.get(prefix + SERVER_PORT_KEY);
    if (serverPortString != null) {
      serverPort = Integer.parseInt(serverPortString);
    }
    String extRules = (String) props.get(prefix + EXTERNAL_RULE_DIRECTORY);
    if (extRules != null) {
      externalRuleDirectory = extRules;
    }

    String paraCheckString = (String) props.get(prefix + NO_DEFAULT_CHECK_KEY);
    if (Boolean.parseBoolean(paraCheckString)) {
      paraCheckString = (String) props.get(prefix + PARA_CHECK_KEY);
      if (paraCheckString != null) {
        numParasToCheck = Integer.parseInt(paraCheckString);
      }
    }

    String colorSelectionString = (String) props.get(prefix + COLOR_SELECTION_KEY);
    if (colorSelectionString != null) {
      colorSelection = Integer.parseInt(colorSelectionString);
    }

    String checkDirectSpeechString = (String) props.get(prefix + CHECK_DIRECT_SPEACH_KEY);
    if (checkDirectSpeechString != null) {
      checkDirectSpeech = Integer.parseInt(checkDirectSpeechString);
    }

    String themeSelectionString = (String) props.get(prefix + THEME_SELECTION_KEY);
    if (themeSelectionString != null) {
      themeSelection = Integer.parseInt(themeSelectionString);
    }

    String resetCheckString = (String) props.get(prefix + RESET_CHECK_KEY);
    if (resetCheckString != null) {
      doResetCheck = Boolean.parseBoolean(resetCheckString);
    }

    String useTextLevelQueueString = (String) props.get(prefix + USE_QUEUE_KEY);
    if (useTextLevelQueueString != null) {
      useTextLevelQueue = Boolean.parseBoolean(useTextLevelQueueString);
    }

    String noBackgroundCheckString = (String) props.get(prefix + NO_BACKGROUND_CHECK_KEY);
    if (noBackgroundCheckString != null) {
      noBackgroundCheck = Boolean.parseBoolean(noBackgroundCheckString);
    }

    String switchOffString = (String) props.get(prefix + LT_SWITCHED_OFF_KEY);
    if (switchOffString != null) {
      switchOff = Boolean.parseBoolean(switchOffString);
    }

    String isMultiThreadString = (String) props.get(prefix + IS_MULTI_THREAD_LO_KEY);
    if (isMultiThreadString != null) {
      isMultiThreadLO = Boolean.parseBoolean(isMultiThreadString);
    }
    
    String doRemoteCheckString = (String) props.get(prefix + DO_REMOTE_CHECK_KEY);
    if (doRemoteCheckString != null) {
      doRemoteCheck = Boolean.parseBoolean(doRemoteCheckString);
    }
    
    String useOtherServerString = (String) props.get(prefix + USE_OTHER_SERVER_KEY);
    if (useOtherServerString != null) {
      useOtherServer = Boolean.parseBoolean(useOtherServerString);
    }
    
    otherServerUrl = (String) props.get(prefix + OTHER_SERVER_URL_KEY);
    if (otherServerUrl != null && !isValidServerUrl(otherServerUrl)) {
      otherServerUrl = null;
    }

    String isPremiumString = (String) props.get(prefix + IS_PREMIUM_KEY);
    if (isPremiumString != null) {
      isPremium = Boolean.parseBoolean(isPremiumString);
    }
    
    remoteUsername = (String) props.get(prefix + REMOTE_USERNAME_KEY);

    remoteApiKey = (String) props.get(prefix + REMOTE_APIKEY_KEY);

    String markSingleCharBoldString = (String) props.get(prefix + MARK_SINGLE_CHAR_BOLD_KEY);
    if (markSingleCharBoldString != null) {
      markSingleCharBold = Boolean.parseBoolean(markSingleCharBoldString);
    }
    
    String useLtSpellCheckerString = (String) props.get(prefix + USE_LT_SPELL_CHECKER_KEY);
    if (useLtSpellCheckerString != null) {
      useLtSpellChecker = Boolean.parseBoolean(useLtSpellCheckerString);
    }
    
    String useLongMessagesString = (String) props.get(prefix + USE_LONG_MESSAGES_KEY);
    if (useLongMessagesString != null) {
      useLongMessages = Boolean.parseBoolean(useLongMessagesString);
    }
    
    String noSynonymsAsSuggestionsString = (String) props.get(prefix + NO_SYNONYMS_AS_SUGGESTIONS_KEY);
    if (noSynonymsAsSuggestionsString != null) {
      noSynonymsAsSuggestions = Boolean.parseBoolean(noSynonymsAsSuggestionsString);
    }
    
    String includeTrackedChangesString = (String) props.get(prefix + INCLUDE_TRACKED_CHANGES_KEY);
    if (includeTrackedChangesString != null) {
      includeTrackedChanges = Boolean.parseBoolean(includeTrackedChangesString);
    }
    
    String enableTmpOffRulesString = (String) props.get(prefix + ENABLE_TMP_OFF_RULES_KEY);
    if (enableTmpOffRulesString != null) {
      enableTmpOffRules = Boolean.parseBoolean(enableTmpOffRulesString);
    }
    
    String enableGoalSpecificRulesString = (String) props.get(prefix + ENABLE_GOAL_SPECIFIC_RULES_KEY);
    if (enableGoalSpecificRulesString != null) {
      enableGoalSpecificRules = Boolean.parseBoolean(enableGoalSpecificRulesString);
    }
    
    String filterOverlappingMatchesString = (String) props.get(prefix + FILTER_OVERLAPPING_MATCHES_KEY);
    if (filterOverlappingMatchesString != null) {
      filterOverlappingMatches = Boolean.parseBoolean(filterOverlappingMatchesString);
    }
    
    String saveLoCacheString = (String) props.get(prefix + SAVE_LO_CACHE_KEY);
    if (saveLoCacheString != null) {
      saveLoCache = Boolean.parseBoolean(saveLoCacheString);
    }
    
    String aiString = (String) props.get(prefix + AI_URL_KEY);
    if (aiString != null) {
      aiUrl = aiString;
    }
    
    aiString = (String) props.get(prefix + AI_MODEL_KEY);
    if (aiString != null) {
      aiModel = aiString;
    }
    
    aiString = (String) props.get(prefix + AI_APIKEY_KEY);
    if (aiString != null) {
      aiApiKey = aiString;
    }
    
    aiString = (String) props.get(prefix + AI_USE_AI_SUPPORT_KEY);
    if (aiString != null) {
      useAiSupport = Boolean.parseBoolean(aiString);
    }
    
    aiString = (String) props.get(prefix + AI_AUTO_CORRECT_KEY);
    if (aiString != null) {
      aiAutoCorrect = Boolean.parseBoolean(aiString);
    }
    
    aiString = (String) props.get(prefix + AI_AUTO_SUGGESTION_KEY);
    if (aiString != null) {
      aiAutoSuggestion = Boolean.parseBoolean(aiString);
    }
    
    aiString = (String) props.get(prefix + AI_SHOW_STYLISTIC_CHANGES_KEY);
    if (aiString != null) {
      aiShowStylisticChanges = Integer.parseInt(aiString);
    }
    
    
    aiString = (String) props.get(prefix + AI_IMG_URL_KEY);
    if (aiString != null) {
      aiImgUrl = aiString;
    }
    
    aiString = (String) props.get(prefix + AI_IMG_MODEL_KEY);
    if (aiString != null) {
      aiImgModel = aiString;
    }
    
    aiString = (String) props.get(prefix + AI_IMG_APIKEY_KEY);
    if (aiString != null) {
      aiImgApiKey = aiString;
    }
    
    aiString = (String) props.get(prefix + AI_USE_AI_IMG_SUPPORT_KEY);
    if (aiString != null) {
      useAiImgSupport = Boolean.parseBoolean(aiString);
    }
    
    aiString = (String) props.get(prefix + AI_TTS_URL_KEY);
    if (aiString != null) {
      aiTtsUrl = aiString;
    }
    
    aiString = (String) props.get(prefix + AI_TTS_MODEL_KEY);
    if (aiString != null) {
      aiTtsModel = aiString;
    }
    
    aiString = (String) props.get(prefix + AI_TTS_APIKEY_KEY);
    if (aiString != null) {
      aiTtsApiKey = aiString;
    }
    
    aiString = (String) props.get(prefix + AI_USE_AI_TTS_SUPPORT_KEY);
    if (aiString != null) {
      useAiTtsSupport = Boolean.parseBoolean(aiString);
    }
    
    String rulesValuesString = (String) props.get(prefix + CONFIGURABLE_RULE_VALUES_KEY + qualifier);
    if (rulesValuesString == null) {
      rulesValuesString = (String) props.get(prefix + CONFIGURABLE_RULE_VALUES_KEY);
    }
    parseConfigurableRuleValues(rulesValuesString);

    String colorsString = (String) props.get(prefix + ERROR_COLORS_KEY);
    parseErrorColors(colorsString);

    String underlineColorsString = (String) props.get(prefix + UNDERLINE_COLORS_KEY);
    parseUnderlineColors(underlineColorsString, underlineColors);

    String underlineRuleColorsString = (String) props.get(prefix + UNDERLINE_RULE_COLORS_KEY);
    parseUnderlineColors(underlineRuleColorsString, underlineRuleColors);

    String underlineDefaultColorsString = (String) props.get(prefix + UNDERLINE_DEFAULT_COLORS_KEY);
    parseUnderlineDefaultColors(underlineDefaultColorsString, underlineDefaultColors);

    String underlineTypesString = (String) props.get(prefix + UNDERLINE_TYPES_KEY);
    parseUnderlineTypes(underlineTypesString, underlineTypes);

    String underlineRulesTypesString = (String) props.get(prefix + UNDERLINE_RULE_TYPES_KEY);
    parseUnderlineTypes(underlineRulesTypesString, underlineRuleTypes);

    //store config for other languages
    loadConfigForOtherLanguages(lang, props, prefix);
  }

  private void parseErrorColors(String colorsString) {
    if (StringUtils.isNotEmpty(colorsString)) {
      String[] typeToColorList = colorsString.split(COLOR_SPLITTER_REGEXP);
      for (String typeToColor : typeToColorList) {
        String[] typeAndColor = typeToColor.split(COLOR_SPLITTER_REGEXP_COLON);
        if (typeAndColor.length != 2) {
          throw new RuntimeException("Could not parse type and color, colon expected: '" + typeToColor + "'");
        }
        ITSIssueType type = ITSIssueType.getIssueType(typeAndColor[0]);
        String hexColor = typeAndColor[1];
        errorColors.put(type, Color.decode(hexColor));
      }
    }
  }

  private void parseUnderlineColors(String colorsString, Map<String, Color> underlineColors) {
    if (StringUtils.isNotEmpty(colorsString)) {
      String[] typeToColorList = colorsString.split(COLOR_SPLITTER_REGEXP);
      for (String typeToColor : typeToColorList) {
        String[] typeAndColor = typeToColor.split(COLOR_SPLITTER_REGEXP_COLON);
        if (typeAndColor.length != 2) {
          throw new RuntimeException("Could not parse type and color, colon expected: '" + typeToColor + "'");
        }
        underlineColors.put(typeAndColor[0], Color.decode(typeAndColor[1]));
      }
    }
  }

  private void parseUnderlineDefaultColors(String colorsString, List<Color> colors) {
    if (StringUtils.isNotEmpty(colorsString)) {
      String[] colorList = colorsString.split(",");
      if (colorList.length != 3) {
        throw new RuntimeException("Could not parse default color, comma expected: '" + colorList + "'");
      }
      colors.clear();
      for (String color : colorList) {
        colors.add(Color.decode(color));
      }
    }
  }

  private void parseUnderlineTypes(String typessString, Map<String, Short> underlineTypes) {
    if (StringUtils.isNotEmpty(typessString)) {
      String[] categoryToTypesList = typessString.split(CONFIGURABLE_RULE_SPLITTER_REGEXP);
      for (String categoryToType : categoryToTypesList) {
        String[] categoryAndType = categoryToType.split(":");
        if (categoryAndType.length != 2) {
          throw new RuntimeException("Could not parse category and type, colon expected: '" + categoryToType + "'");
        }
        underlineTypes.put(categoryAndType[0], Short.parseShort(categoryAndType[1]));
      }
    }
  }

  private void parseConfigurableRuleValues(String rulesValueString) {
    if (StringUtils.isNotEmpty(rulesValueString)) {
      String[] ruleToValueList = rulesValueString.split(",");
      for (String ruleToValue : ruleToValueList) {
        ruleToValue = ruleToValue.trim();
        if (!ruleToValue.isEmpty()) {
          String[] ruleAndValue = ruleToValue.split(":");
          if (ruleAndValue.length != 2) {
            throw new RuntimeException("Could not parse rule and value, colon expected: '" + ruleToValue + "'");
          }
          Object[] objects = RuleOption.stringToObjects(ruleAndValue[1]);
          configurableRuleValues.put(ruleAndValue[0].trim(), objects);
        }
      }
    }
  }

  private String getQualifier(Language lang) {
    String qualifier = "";
    if (lang != null) {
      qualifier = "." + lang.getShortCodeWithCountryAndVariant();
    }
    return qualifier;
  }

  private void loadConfigForOtherLanguages(Language lang, Properties prop, String prefix) {
    for (Language otherLang : Languages.get()) {
      if (!otherLang.equals(lang)) {
        String languageSuffix = "." + otherLang.getShortCodeWithCountryAndVariant();
        storeConfigKeyFromProp(prop, prefix + DISABLED_RULES_KEY + languageSuffix);
        storeConfigKeyFromProp(prop, prefix + ENABLED_RULES_KEY + languageSuffix);
        storeConfigKeyFromProp(prop, prefix + DISABLED_CATEGORIES_KEY + languageSuffix);
        storeConfigKeyFromProp(prop, prefix + ENABLED_CATEGORIES_KEY + languageSuffix);
        storeConfigKeyFromProp(prop, prefix + CONFIGURABLE_RULE_VALUES_KEY + languageSuffix);
      }
    }
  }

  private void storeConfigKeyFromProp(Properties prop, String key) {
    if (prop.containsKey(key)) {
      configForOtherLanguages.put(key, prop.getProperty(key));
    }
  }

  private Collection<? extends String> getListFromProperties(Properties props, String key) {
    String value = (String) props.get(key);
    List<String> list = new ArrayList<>();
    if (value != null && !value.isEmpty()) {
      String[] names = value.split(DELIMITER);
      for (int i = 0; i < names.length; i++) {
        names[i] = names[i].replace(COMMA_REPLACE, DELIMITER);
//        WtMessageHandler.printToLogFile("new Name: " + name);
      }
      list.addAll(Arrays.asList(names));
    }
    return list;
  }

  public void saveConfiguration(Language lang) throws IOException {
    if (lang == null) {
      lang = this.lang;
    }
    Properties props = new Properties();
    String qualifier = getQualifier(lang);

    String[] versionParts = WtVersionInfo.wtVersion.split("-");
    props.setProperty(LT_VERSION_KEY, versionParts[0]);

    if (currentProfile != null && !currentProfile.isEmpty()) {
      props.setProperty(CURRENT_PROFILE_KEY, currentProfile);
    }
    
    if (!definedProfiles.isEmpty()) {
      props.setProperty(DEFINED_PROFILES_KEY, String.join(DELIMITER, definedProfiles));
    }
    
    if (motherTongue != null) {
      props.setProperty(MOTHER_TONGUE_KEY, motherTongue.getShortCodeWithCountryAndVariant());
    }

    if (logLevel != null) {
      props.setProperty(LOG_LEVEL_KEY, logLevel);
    }
    
    try (FileOutputStream fos = new FileOutputStream(configFile)) {
      props.store(fos, WtOfficeTools.WT_NAME + " configuration (" + WtVersionInfo.getWtInformation() + ")");
    }

    List<String> prefixes = new ArrayList<>();
    prefixes.add("");
    for (String profile : definedProfiles) {
      String prefix = profile;
      prefixes.add(prefix.replaceAll(BLANK, BLANK_REPLACE) + PROFILE_DELIMITER);
    }
    String currentPrefix;
    if (currentProfile == null) {
      currentPrefix = "";
    } else {
      currentPrefix = currentProfile;
    }
    if (!currentPrefix.isEmpty()) {
      currentPrefix = currentPrefix.replaceAll(BLANK, BLANK_REPLACE);
      currentPrefix += PROFILE_DELIMITER;
    }
    for (String prefix : prefixes) {
      props = new Properties();
      if (currentPrefix.equals(prefix)) {
        saveConfigForCurrentProfile(props, prefix, qualifier);
      } else {
        saveConfigForProfile(props, prefix);
      }

      try (FileOutputStream fos = new FileOutputStream(configFile, true)) {
        props.store(fos, "Profile: " + (prefix.isEmpty() ? "Default" : prefix.substring(0, prefix.length() - 2)));
      }
    }
    
  }

  private void addListToProperties(Properties props, String key, Set<String> list) {
    if (list == null) {
      props.setProperty(key, "");
    } else {
      Set<String> corList = new HashSet<>();
      for (String entry : list) {
        corList.add(entry.replace(DELIMITER, COMMA_REPLACE));
      }
      props.setProperty(key, String.join(DELIMITER, corList));
    }
  }
  
  private void setAllProfileKeys() {
    allProfileKeys.add(LANGUAGE_KEY);
    allProfileKeys.add(FIXED_LANGUAGE_KEY);
    allProfileKeys.add(NGRAM_DIR_KEY);
    allProfileKeys.add(AUTO_DETECT_KEY);
    allProfileKeys.add(TAGGER_SHOWS_DISAMBIG_LOG_KEY);
    allProfileKeys.add(SERVER_RUN_KEY);
    allProfileKeys.add(SERVER_PORT_KEY);
    allProfileKeys.add(NO_DEFAULT_CHECK_KEY);
    allProfileKeys.add(PARA_CHECK_KEY);
    allProfileKeys.add(COLOR_SELECTION_KEY);
    allProfileKeys.add(CHECK_DIRECT_SPEACH_KEY);
    allProfileKeys.add(THEME_SELECTION_KEY);
    allProfileKeys.add(RESET_CHECK_KEY);
    allProfileKeys.add(USE_QUEUE_KEY);
    allProfileKeys.add(NO_BACKGROUND_CHECK_KEY);
    allProfileKeys.add(USE_DOC_LANG_KEY);
    allProfileKeys.add(USE_GUI_KEY);
    allProfileKeys.add(FONT_NAME_KEY);
    allProfileKeys.add(FONT_STYLE_KEY);
    allProfileKeys.add(FONT_SIZE_KEY);
    allProfileKeys.add(LF_NAME_KEY);
    allProfileKeys.add(ERROR_COLORS_KEY);
    allProfileKeys.add(UNDERLINE_COLORS_KEY);
    allProfileKeys.add(UNDERLINE_RULE_COLORS_KEY);
    allProfileKeys.add(UNDERLINE_DEFAULT_COLORS_KEY);
    allProfileKeys.add(UNDERLINE_TYPES_KEY);
    allProfileKeys.add(UNDERLINE_RULE_TYPES_KEY);
    allProfileKeys.add(LT_SWITCHED_OFF_KEY);
    allProfileKeys.add(IS_MULTI_THREAD_LO_KEY);
    allProfileKeys.add(EXTERNAL_RULE_DIRECTORY);
    allProfileKeys.add(DO_REMOTE_CHECK_KEY);
    allProfileKeys.add(OTHER_SERVER_URL_KEY);
    allProfileKeys.add(USE_OTHER_SERVER_KEY);
    allProfileKeys.add(IS_PREMIUM_KEY);
    allProfileKeys.add(REMOTE_USERNAME_KEY);
    allProfileKeys.add(REMOTE_APIKEY_KEY);
    allProfileKeys.add(MARK_SINGLE_CHAR_BOLD_KEY);
    allProfileKeys.add(USE_LT_SPELL_CHECKER_KEY);
    allProfileKeys.add(USE_LONG_MESSAGES_KEY);
    allProfileKeys.add(NO_SYNONYMS_AS_SUGGESTIONS_KEY);
    allProfileKeys.add(INCLUDE_TRACKED_CHANGES_KEY);
    allProfileKeys.add(ENABLE_TMP_OFF_RULES_KEY);
    allProfileKeys.add(ENABLE_GOAL_SPECIFIC_RULES_KEY);
    allProfileKeys.add(FILTER_OVERLAPPING_MATCHES_KEY);
    allProfileKeys.add(SAVE_LO_CACHE_KEY);
    allProfileKeys.add(AI_URL_KEY);
    allProfileKeys.add(AI_APIKEY_KEY);
    allProfileKeys.add(AI_MODEL_KEY);
    allProfileKeys.add(AI_USE_AI_SUPPORT_KEY);
    allProfileKeys.add(AI_AUTO_CORRECT_KEY);
    allProfileKeys.add(AI_AUTO_SUGGESTION_KEY);
    allProfileKeys.add(AI_SHOW_STYLISTIC_CHANGES_KEY);
    allProfileKeys.add(AI_IMG_URL_KEY);
    allProfileKeys.add(AI_IMG_APIKEY_KEY);
    allProfileKeys.add(AI_IMG_MODEL_KEY);
    allProfileKeys.add(AI_USE_AI_IMG_SUPPORT_KEY);
    allProfileKeys.add(AI_TTS_URL_KEY);
    allProfileKeys.add(AI_TTS_APIKEY_KEY);
    allProfileKeys.add(AI_TTS_MODEL_KEY);
    allProfileKeys.add(AI_USE_AI_TTS_SUPPORT_KEY);

    allProfileLangKeys.add(DISABLED_RULES_KEY);
    allProfileLangKeys.add(ENABLED_RULES_KEY);
    allProfileLangKeys.add(DISABLED_CATEGORIES_KEY);
    allProfileLangKeys.add(ENABLED_CATEGORIES_KEY);
    allProfileLangKeys.add(CONFIGURABLE_RULE_VALUES_KEY);
  }
  
  private void storeConfigForAllProfiles(Properties props) {
    List<String> prefix = new ArrayList<>();
    prefix.add("");
    for (String profile : definedProfiles) {
      String sPrefix = profile;
      prefix.add(sPrefix.replaceAll(BLANK, BLANK_REPLACE) + PROFILE_DELIMITER);
    }
    for (String sPrefix : prefix) {
      for (String key : allProfileLangKeys) {
        for (Language lang : Languages.get()) {
          String preKey = sPrefix + key + "." + lang.getShortCodeWithCountryAndVariant();
          if (props.containsKey(preKey)) {
            configForOtherProfiles.put(preKey, props.getProperty(preKey));
          }
        }
      }
    }
    for (String sPrefix : prefix) {
      if (isOffice && ltVersion == null) {
        if (props.containsKey(sPrefix + MOTHER_TONGUE_KEY)) {
          configForOtherProfiles.put(sPrefix + FIXED_LANGUAGE_KEY, props.getProperty(sPrefix + MOTHER_TONGUE_KEY));
        }
      }
      for (String key : allProfileKeys) {
        String preKey = sPrefix + key;
        if (props.containsKey(preKey)) {
          configForOtherProfiles.put(preKey, props.getProperty(preKey));
        }
      }
    }
  }
  
  private void saveConfigForCurrentProfile(Properties props, String prefix, String qualifier) {
    if (!disabledRuleIds.isEmpty()) {
      addListToProperties(props, prefix + DISABLED_RULES_KEY + qualifier, disabledRuleIds);
    }
    if (!enabledRuleIds.isEmpty()) {
      addListToProperties(props, prefix + ENABLED_RULES_KEY + qualifier, enabledRuleIds);
    }
    if (!disabledCategoryNames.isEmpty()) {
      addListToProperties(props, prefix + DISABLED_CATEGORIES_KEY + qualifier, disabledCategoryNames);
    }
    if (!enabledCategoryNames.isEmpty()) {
      addListToProperties(props, prefix + ENABLED_CATEGORIES_KEY + qualifier, enabledCategoryNames);
    }
    if (language != null && !language.isExternal()) {  // external languages won't be known at startup, so don't save them
      props.setProperty(prefix + LANGUAGE_KEY, language.getShortCodeWithCountryAndVariant());
    }
    if (fixedLanguage != null) {
      props.setProperty(prefix + FIXED_LANGUAGE_KEY, fixedLanguage.getShortCodeWithCountryAndVariant());
    }
    if (ngramDirectory != null) {
      props.setProperty(prefix + NGRAM_DIR_KEY, ngramDirectory.getAbsolutePath());
    }
    props.setProperty(prefix + AUTO_DETECT_KEY, Boolean.toString(autoDetect));
    props.setProperty(prefix + TAGGER_SHOWS_DISAMBIG_LOG_KEY, Boolean.toString(taggerShowsDisambigLog));
    props.setProperty(prefix + USE_GUI_KEY, Boolean.toString(guiConfig));
    props.setProperty(prefix + SERVER_RUN_KEY, Boolean.toString(runServer));
    props.setProperty(prefix + SERVER_PORT_KEY, Integer.toString(serverPort));
    if (numParasToCheck != DEFAULT_NUM_CHECK_PARAS) {
      props.setProperty(prefix + NO_DEFAULT_CHECK_KEY, Boolean.toString(true));
      props.setProperty(prefix + PARA_CHECK_KEY, Integer.toString(numParasToCheck));
    }
    if (colorSelection != DEFAULT_COLOR_SELECTION) {
      props.setProperty(prefix + COLOR_SELECTION_KEY, Integer.toString(colorSelection));
    }
    if (checkDirectSpeech != DEFAULT_CHECK_DIRECT_SPEECH) {
      props.setProperty(prefix + CHECK_DIRECT_SPEACH_KEY, Integer.toString(checkDirectSpeech));
    }
    if (themeSelection != DEFAULT_THEME_SELECTION) {
      props.setProperty(prefix + THEME_SELECTION_KEY, Integer.toString(themeSelection));
    }
    if (doResetCheck != DEFAULT_DO_RESET) {
      props.setProperty(prefix + RESET_CHECK_KEY, Boolean.toString(doResetCheck));
    }
    if (useTextLevelQueue != DEFAULT_USE_QUEUE) {
      props.setProperty(prefix + USE_QUEUE_KEY, Boolean.toString(useTextLevelQueue));
    }
    if (noBackgroundCheck != DEFAULT_NO_BACKGROUND_CHECK) {
      props.setProperty(prefix + NO_BACKGROUND_CHECK_KEY, Boolean.toString(noBackgroundCheck));
    }
    if (useDocLanguage != DEFAULT_USE_DOC_LANGUAGE) {
      props.setProperty(prefix + USE_DOC_LANG_KEY, Boolean.toString(useDocLanguage));
    }
    if (isMultiThreadLO != DEFAULT_MULTI_THREAD) {
      props.setProperty(prefix + IS_MULTI_THREAD_LO_KEY, Boolean.toString(isMultiThreadLO));
    }
    if (doRemoteCheck != DEFAULT_DO_REMOTE_CHECK) {
      props.setProperty(prefix + DO_REMOTE_CHECK_KEY, Boolean.toString(doRemoteCheck));
    }
    if (useOtherServer != DEFAULT_USE_OTHER_SERVER) {
      props.setProperty(prefix + USE_OTHER_SERVER_KEY, Boolean.toString(useOtherServer));
    }
    if (isPremium != DEFAULT_IS_PREMIUM) {
      props.setProperty(prefix + IS_PREMIUM_KEY, Boolean.toString(isPremium));
    }
    if (markSingleCharBold != DEFAULT_MARK_SINGLE_CHAR_BOLD) {
      props.setProperty(prefix + MARK_SINGLE_CHAR_BOLD_KEY, Boolean.toString(markSingleCharBold));
    }
    if (useLtSpellChecker != DEFAULT_USE_LT_SPELL_CHECKER) {
      props.setProperty(prefix + USE_LT_SPELL_CHECKER_KEY, Boolean.toString(useLtSpellChecker));
    }
    if (useLongMessages != DEFAULT_USE_LONG_MESSAGES) {
      props.setProperty(prefix + USE_LONG_MESSAGES_KEY, Boolean.toString(useLongMessages));
    }
    if (noSynonymsAsSuggestions != DEFAULT_NO_SYNONYMS_AS_SUGGESTIONS) {
      props.setProperty(prefix + NO_SYNONYMS_AS_SUGGESTIONS_KEY, Boolean.toString(noSynonymsAsSuggestions));
    }
    if (includeTrackedChanges != DEFAULT_INCLUDE_TRACKED_CHANGES) {
      props.setProperty(prefix + INCLUDE_TRACKED_CHANGES_KEY, Boolean.toString(includeTrackedChanges));
    }
    if (enableTmpOffRules != DEFAULT_ENABLE_TMP_OFF_RULES) {
      props.setProperty(prefix + ENABLE_TMP_OFF_RULES_KEY, Boolean.toString(enableTmpOffRules));
    }
    if (enableGoalSpecificRules != DEFAULT_ENABLE_GOAL_SPECIFIC_RULES) {
      props.setProperty(prefix + ENABLE_GOAL_SPECIFIC_RULES_KEY, Boolean.toString(enableGoalSpecificRules));
    }
    if (filterOverlappingMatches != DEFAULT_FILTER_OVERLAPPING_MATCHES) {
      props.setProperty(prefix + FILTER_OVERLAPPING_MATCHES_KEY, Boolean.toString(filterOverlappingMatches));
    }
    if (saveLoCache != DEFAULT_SAVE_LO_CACHE) {
      props.setProperty(prefix + SAVE_LO_CACHE_KEY, Boolean.toString(saveLoCache));
    }
    if (switchOff) {
      props.setProperty(prefix + LT_SWITCHED_OFF_KEY, Boolean.toString(switchOff));
    }
    if (otherServerUrl != null && isValidServerUrl(otherServerUrl)) {
      props.setProperty(prefix + OTHER_SERVER_URL_KEY, otherServerUrl);
    }
    if (remoteUsername != null) {
      props.setProperty(prefix + REMOTE_USERNAME_KEY, remoteUsername);
    }
    if (remoteApiKey != null) {
      props.setProperty(prefix + REMOTE_APIKEY_KEY, remoteApiKey);
    }
    if (aiUrl != null) {
      props.setProperty(prefix + AI_URL_KEY, aiUrl);
    }
    if (aiModel != null) {
      props.setProperty(prefix + AI_MODEL_KEY, aiModel);
    }
    if (aiApiKey != null) {
      props.setProperty(prefix + AI_APIKEY_KEY, aiApiKey);
    }
    if (useAiSupport != DEFAULT_USE_AI_SUPPORT) {
      props.setProperty(prefix + AI_USE_AI_SUPPORT_KEY, Boolean.toString(useAiSupport));
    }
    if (aiAutoCorrect != DEFAULT_AI_AUTO_CORRECT) {
      props.setProperty(prefix + AI_AUTO_CORRECT_KEY, Boolean.toString(aiAutoCorrect));
    }
    if (aiAutoSuggestion != DEFAULT_AI_AUTO_SUGGESTION) {
      props.setProperty(prefix + AI_AUTO_SUGGESTION_KEY, Boolean.toString(aiAutoSuggestion));
    }
    if (this.aiShowStylisticChanges != DEFAULT_AI_SHOW_STYLISTIC_CHANGES) {
      props.setProperty(prefix + AI_SHOW_STYLISTIC_CHANGES_KEY, Integer.toString(aiShowStylisticChanges));
    }
    if (aiImgUrl != null) {
      props.setProperty(prefix + AI_IMG_URL_KEY, aiImgUrl);
    }
    if (aiImgModel != null) {
      props.setProperty(prefix + AI_IMG_MODEL_KEY, aiImgModel);
    }
    if (aiImgApiKey != null) {
      props.setProperty(prefix + AI_IMG_APIKEY_KEY, aiImgApiKey);
    }
    if (useAiImgSupport != DEFAULT_USE_AI_IMG_SUPPORT) {
      props.setProperty(prefix + AI_USE_AI_IMG_SUPPORT_KEY, Boolean.toString(useAiImgSupport));
    }
    if (aiTtsUrl != null) {
      props.setProperty(prefix + AI_TTS_URL_KEY, aiTtsUrl);
    }
    if (aiTtsModel != null) {
      props.setProperty(prefix + AI_TTS_MODEL_KEY, aiTtsModel);
    }
    if (aiTtsApiKey != null) {
      props.setProperty(prefix + AI_TTS_APIKEY_KEY, aiTtsApiKey);
    }
    if (useAiTtsSupport != DEFAULT_USE_AI_TTS_SUPPORT) {
      props.setProperty(prefix + AI_USE_AI_TTS_SUPPORT_KEY, Boolean.toString(useAiTtsSupport));
    }
    if (fontName != null) {
      props.setProperty(prefix + FONT_NAME_KEY, fontName);
    }
    if (fontStyle != FONT_STYLE_INVALID) {
      props.setProperty(prefix + FONT_STYLE_KEY, Integer.toString(fontStyle));
    }
    if (fontSize != FONT_SIZE_INVALID) {
      props.setProperty(prefix + FONT_SIZE_KEY, Integer.toString(fontSize));
    }
    if (this.lookAndFeelName != null) {
      props.setProperty(prefix + LF_NAME_KEY, lookAndFeelName);
    }
    if (externalRuleDirectory != null) {
      props.setProperty(prefix + EXTERNAL_RULE_DIRECTORY, externalRuleDirectory);
    }
    if (!configurableRuleValues.isEmpty()) {
      StringBuilder sbRV = new StringBuilder();
      int i = 0;
      for (Map.Entry<String, Object[]> entry : configurableRuleValues.entrySet()) {
        Object[] obs = entry.getValue();
        if (obs != null && obs.length > 0) {
          if (i > 0) {
            sbRV.append(",");
          }
          sbRV.append(entry.getKey()).append(':').append(RuleOption.objectsToString(obs));
          i++;
        }
      }
      props.setProperty(prefix + CONFIGURABLE_RULE_VALUES_KEY + qualifier, sbRV.toString());
    }
    if (!errorColors.isEmpty()) {
      StringBuilder sb = new StringBuilder();
      for (Map.Entry<ITSIssueType, Color> entry : errorColors.entrySet()) {
        String rgb = Integer.toHexString(entry.getValue().getRGB());
        rgb = rgb.substring(2);
        sb.append(entry.getKey()).append(":#").append(rgb).append(", ");
      }
      props.setProperty(prefix + ERROR_COLORS_KEY, sb.toString());
    }
    if (!underlineColors.isEmpty()) {
      StringBuilder sbUC = new StringBuilder();
      for (Map.Entry<String, Color> entry : underlineColors.entrySet()) {
        String rgb = Integer.toHexString(entry.getValue().getRGB());
        rgb = rgb.substring(2);
        sbUC.append(entry.getKey()).append(":#").append(rgb).append(", ");
      }
      props.setProperty(prefix + UNDERLINE_COLORS_KEY, sbUC.toString());
    }
    if (!underlineRuleColors.isEmpty()) {
      StringBuilder sbUC = new StringBuilder();
      for (Map.Entry<String, Color> entry : underlineRuleColors.entrySet()) {
        String rgb = Integer.toHexString(entry.getValue().getRGB());
        rgb = rgb.substring(2);
        sbUC.append(entry.getKey()).append(":#").append(rgb).append(", ");
      }
      props.setProperty(prefix + UNDERLINE_RULE_COLORS_KEY, sbUC.toString());
    }
    if (!underlineDefaultColors.isEmpty()) {
      StringBuilder sbUC = new StringBuilder();
      int i = 0;
      for (Color entry : underlineDefaultColors) {
        String rgb = Integer.toHexString(entry.getRGB());
        rgb = rgb.substring(2);
        if (i > 0) {
          sbUC = sbUC.append(",");
        }
        sbUC.append("#").append(rgb);
        i++;
      }
      props.setProperty(prefix + UNDERLINE_DEFAULT_COLORS_KEY, sbUC.toString());
    }
    if (!underlineTypes.isEmpty()) {
      StringBuilder sbUT = new StringBuilder();
      for (Map.Entry<String, Short> entry : underlineTypes.entrySet()) {
        sbUT.append(entry.getKey()).append(':').append(entry.getValue()).append(", ");
      }
      props.setProperty(prefix + UNDERLINE_TYPES_KEY, sbUT.toString());
    }
    if (!underlineRuleTypes.isEmpty()) {
      StringBuilder sbUT = new StringBuilder();
      for (Map.Entry<String, Short> entry : underlineRuleTypes.entrySet()) {
        sbUT.append(entry.getKey()).append(':').append(entry.getValue()).append(", ");
      }
      props.setProperty(prefix + UNDERLINE_RULE_TYPES_KEY, sbUT.toString());
    }
    for (String key : configForOtherLanguages.keySet()) {
      props.setProperty(key, configForOtherLanguages.get(key));
    }
  }

  private void saveConfigForProfile(Properties props, String prefix) {
    for (String key : allProfileLangKeys) {
      for (Language lang : Languages.get()) {
        String preKey = prefix + key + "." + lang.getShortCodeWithCountryAndVariant();
        if (configForOtherProfiles.containsKey(preKey)) {
          props.setProperty(preKey, configForOtherProfiles.get(preKey));
        }
      }
    }
    for (String key : allProfileKeys) {
      String preKey = prefix + key;
      if (configForOtherProfiles.containsKey(preKey)) {
        props.setProperty(preKey, configForOtherProfiles.get(preKey));
      }
    }
  }

  public void importProfile(File importFile) throws IOException {
    String qualifier = getQualifier(lang);
    try (FileInputStream fis = new FileInputStream(importFile)) {
      Properties props = new Properties();
      props.load(fis);
      String curProfileStr = (String) props.get(CURRENT_PROFILE_KEY);
      if (curProfileStr == null || curProfileStr.isEmpty()) {
        return;
      }
      currentProfile = curProfileStr;
      String prefix = currentProfile;
      prefix = prefix.replaceAll(BLANK, BLANK_REPLACE);
      prefix += PROFILE_DELIMITER;
      loadCurrentProfile(props, prefix, qualifier);
    } catch (FileNotFoundException e) {
      // file not found: okay, leave disabledRuleIds empty
    }
  }

  public void exportProfile(String profile, File exportFile) throws IOException {
    Properties props = new Properties();
    String qualifier = getQualifier(lang);
    if (currentProfile != null && !currentProfile.isEmpty()) {
      props.setProperty(CURRENT_PROFILE_KEY, profile);
    }
    try (FileOutputStream fos = new FileOutputStream(exportFile)) {
      props.store(fos, WtOfficeTools.WT_NAME + " configuration (" + WtVersionInfo.getWtInformation() + ")");
    }
    String prefix = profile;
    if (!prefix.isEmpty()) {
      prefix = prefix.replaceAll(BLANK, BLANK_REPLACE);
      prefix += PROFILE_DELIMITER;
    }
    saveConfigForCurrentProfile(props, prefix, qualifier);
    try (FileOutputStream fos = new FileOutputStream(exportFile, true)) {
      props.store(fos, "Profile: " + (prefix.isEmpty() ? "Default" : prefix.substring(0, prefix.length() - 2)));
    }
  }

  
}
