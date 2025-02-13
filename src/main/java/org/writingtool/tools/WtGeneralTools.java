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

import org.apache.commons.lang3.StringUtils;
import org.languagetool.JLanguageTool;
import org.languagetool.rules.*;
import org.languagetool.rules.patterns.FalseFriendPatternRule;
import org.languagetool.tools.StringTools;
import org.writingtool.WtDocumentsHandler;
import org.writingtool.config.WtConfiguration;
import org.writingtool.dialogs.WtOptionPane;

import com.formdev.flatlaf.FlatDarkLaf;
import com.formdev.flatlaf.FlatLightLaf;

import javax.swing.*;
import javax.swing.UIManager.LookAndFeelInfo;
import javax.swing.event.HyperlinkEvent;
import javax.swing.event.HyperlinkListener;
import javax.swing.filechooser.FileFilter;
import java.awt.*;
import java.io.File;
import java.io.UnsupportedEncodingException;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLEncoder;
import java.util.List;
import java.util.ResourceBundle;
import java.util.Set;
import org.apache.commons.lang3.SystemUtils;

/**
 * GUI-related tools.
 * 
 * @author Fred Kruse
 */
public final class WtGeneralTools {
  
  public final static int THEME_SYSTEM = 0;
  public final static int THEME_FLATDARK = 1;
  public final static int THEME_FLATLIGHT = 2;

  private WtGeneralTools() {
    // no public constructor
  }

  /**
   * Show a file chooser dialog and return the file selected by the user or
   * <code>null</code>.
   */
  public static File openFileDialog(Frame frame, FileFilter fileFilter) {
    return openFileDialog(frame, fileFilter, null);
  }

  /**
   * Show a file chooser dialog in a specified directory
   * @param frame Owner frame
   * @param fileFilter The pattern of files to choose from
   * @param initialDir The initial directory
   * @return the selected file
   * @since 2.6
   */
  public static File openFileDialog(Frame frame, FileFilter fileFilter, File initialDir) {
    return openFileDialog(frame, fileFilter, initialDir, JFileChooser.FILES_ONLY);
  }

  /**
   * Show a directory chooser dialog, starting with a specified directory
   * @param frame Owner frame
   * @param initialDir The initial directory
   * @return the selected file
   * @since 3.0
   */
  public static File openDirectoryDialog(Frame frame, File initialDir) {
    return openFileDialog(frame, null, initialDir, JFileChooser.DIRECTORIES_ONLY);
  }

  private static File openFileDialog(Frame frame, FileFilter fileFilter, File initialDir, int mode) {
    JFileChooser jfc = new JFileChooser();
    jfc.setFileSelectionMode(mode);
    jfc.setCurrentDirectory(initialDir);
    jfc.setFileFilter(fileFilter);
    jfc.showOpenDialog(frame);
    return jfc.getSelectedFile();
  }

  /**
   * Show the exception (with stacktrace) in a dialog and print it to STDERR.
   */
  public static void showError(Exception e) {
    String stackTrace = org.languagetool.tools.Tools.getFullStackTrace(e);
    String msg = "<html><p style='width: 600px;'>" + StringTools.escapeHTML(stackTrace);
    WtOptionPane.showMessageDialog(null, msg, "Error", WtOptionPane.ERROR_MESSAGE);
    e.printStackTrace();
  }

  /**
   * Show the exception (message without stacktrace) in a dialog and print the
   * stacktrace to STDERR.
   */
  public static void showErrorMessage(Exception e, Component parent) {
    String msg = e.getMessage();
    WtOptionPane.showMessageDialog(parent, msg, "Error", WtOptionPane.ERROR_MESSAGE);
    e.printStackTrace();
  }

  /**
   * Show the exception (message without stacktrace) in a dialog and print the
   * stacktrace to STDERR.
   */
  public static void showErrorMessage(Exception e) {
    showErrorMessage(e, null);
  }

  /**
   * LibO shortens menu items with more than ~100 characters by dropping text in the middle.
   * That isn't really sensible, so we shorten the text here in order to preserve the important parts.
   */
  public static String shortenComment(String comment) {
    int maxCommentLength = 100;
    String shortComment = comment;
    if (shortComment.length() > maxCommentLength) {
      // if there is text in brackets, drop it (beginning at the end)
      while (shortComment.lastIndexOf(" [") > 0
              && shortComment.lastIndexOf(']') > shortComment.lastIndexOf(" [")
              && shortComment.length() > maxCommentLength) {
        shortComment = shortComment.substring(0, shortComment.lastIndexOf(" [")) + shortComment.substring(shortComment.lastIndexOf(']')+1);
      }
      while (shortComment.lastIndexOf(" (") > 0
              && shortComment.lastIndexOf(')') > shortComment.lastIndexOf(" (")
              && shortComment.length() > maxCommentLength) {
        shortComment = shortComment.substring(0, shortComment.lastIndexOf(" (")) + shortComment.substring(shortComment.lastIndexOf(')')+1);
      }
      // in case it's still not short enough, shorten at the end
      if (shortComment.length() > maxCommentLength) {
        shortComment = shortComment.substring(0, maxCommentLength-1) + "…";
      }
    }
    return shortComment;
  }

  /**
   * Returns translation of the UI element without the control character {@code &}. To
   * have {@code &} in the UI, use {@code &&}.
   *
   * @param label Label to convert.
   * @return String UI element string without mnemonics.
   */
  public static String getLabel(String label) {
    return label.replaceAll("&([^&])", "$1").replaceAll("&&", "&");
  }

  /**
   * Returns mnemonic of a UI element.
   *
   * @param label String Label of the UI element
   * @return Mnemonic of the UI element, or {@code \u0000} in case of no mnemonic set.
   */
  public static char getMnemonic(String label) {
    int mnemonicPos = label.indexOf('&');
    while (mnemonicPos != -1 && mnemonicPos == label.indexOf("&&")
            && mnemonicPos < label.length()) {
      mnemonicPos = label.indexOf('&', mnemonicPos + 2);
    }
    if (mnemonicPos == -1 || mnemonicPos == label.length()) {
      return '\u0000';
    }
    return label.charAt(mnemonicPos + 1);
  }

  /**
   * Set dialog location to the center of the screen
   *
   * @param dialog the dialog which will be centered
   * @since 2.6
   */
  public static void centerDialog(JDialog dialog) {
    Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
    Dimension frameSize = dialog.getSize();
    dialog.setLocation(screenSize.width / 2 - frameSize.width / 2,
            screenSize.height / 2 - frameSize.height / 2);
    dialog.setLocationByPlatform(true);
  }

  /**
   * @since 3.3
   */
  public static void configureFromRules(JLanguageTool lt, WtConfiguration config) {
    Set<String> disabledRules = config.getDisabledRuleIds();
    if (disabledRules != null) {
      for (String ruleId : disabledRules) {
        lt.disableRule(ruleId);
      }
    }
    Set<String> disabledCategories = config.getDisabledCategoryNames();
    if (disabledCategories != null) {
      for (String categoryName : disabledCategories) {
        lt.disableCategory(new CategoryId(categoryName));
      }
    }
    if(config.getEnabledRulesOnly()) {
      for (Rule rule : lt.getAllRules()) {
        lt.disableRule(rule.getId());
      }
    }
    Set<String> enabledRules = config.getEnabledRuleIds();
    if (enabledRules != null) {
      for (String ruleName : enabledRules) {
        lt.enableRule(ruleName);
      }
    }
    lt.setConfigValues(config.getConfigurableValues());
  }
  
  public static void addHyperlinkListener(JTextPane pane) {
    pane.addHyperlinkListener(new HyperlinkListener() {
      @Override
      public void hyperlinkUpdate(HyperlinkEvent e) {
        if (e.getEventType() == HyperlinkEvent.EventType.ACTIVATED) {
          WtGeneralTools.openURL(e.getURL());
        }
      }
    });
  }

  /**
   * Launches the default browser to display a URL.
   * 
   * @param url the URL to be displayed
   * @since 4.1
   */
  public static void openURL(String url) {
    try {
      openURL(new URL(url));
    } catch (MalformedURLException ex) {
      WtGeneralTools.showError(ex);
    }
  }

  /**
   * Launches the default browser to display a URL.
   * 
   * @param url the URL to be displayed
   * @since 4.1
   */
  public static void openURL(URL url) {
    if (Desktop.isDesktopSupported() 
        && Desktop.getDesktop().isSupported(Desktop.Action.BROWSE)) {
      try {
        Desktop.getDesktop().browse(url.toURI());
      } catch (Exception ex) {
        WtGeneralTools.showError(ex);
      }
    } else if(SystemUtils.IS_OS_LINUX) {
      //handle the case where Desktop.browse() is not supported, e.g. kubuntu
      //without libgnome
      try {
        Runtime.getRuntime().exec(new String[] { "xdg-open", url.toString() });
      } catch (Exception ex) {
        WtGeneralTools.showError(ex);
      }
    }
  }

  public static void showRuleInfoDialog(Component parent, String title, String message, Rule rule, URL matchUrl, ResourceBundle messages, String lang) {
    try {
//      int theme = WtDocumentsHandler.getJavaLookAndFeelSet();
//      WtGeneralTools.setJavaLookAndFeel(WtGeneralTools.THEME_SYSTEM);

      int dialogWidth = 320;
      JTextPane textPane = new JTextPane();
      textPane.setEditable(false);
      textPane.setContentType("text/html");
      textPane.setBorder(BorderFactory.createEmptyBorder());
      textPane.setOpaque(false);
      textPane.setBackground(new Color(0, 0, 0, 0));
      WtGeneralTools.addHyperlinkListener(textPane);
      textPane.setSize(dialogWidth, Short.MAX_VALUE);
      String messageWithBold = message.replaceAll("<suggestion>", "<b>").replaceAll("</suggestion>", "</b>");
      String exampleSentences = getExampleSentences(rule, messages);
      String url = "http://community.languagetool.org/rule/show/" + encodeUrl(rule)
              + "?lang=" + lang + "&amp;ref=standalone-gui";
      boolean isExternal = rule.getCategory().getLocation() == Category.Location.EXTERNAL;
      String ruleDetailLink = rule instanceof FalseFriendPatternRule || isExternal ?
              "" : "<a href='" + url + "'>" + messages.getString("ruleDetailsLink") +"</a>";
      textPane.setText("<html>"
              + messageWithBold + exampleSentences + formatURL(matchUrl)
              + "<br><br>"
              + ruleDetailLink
              + "</html>");
      JScrollPane scrollPane = new JScrollPane(textPane);
      scrollPane.setPreferredSize(
              new Dimension(dialogWidth, textPane.getPreferredSize().height));
      scrollPane.setBorder(BorderFactory.createEmptyBorder());
  
      String cleanTitle = title.replace("<suggestion>", "'").replace("</suggestion>", "'");
      JOptionPane.showMessageDialog(parent, scrollPane, cleanTitle,
              WtOptionPane.INFORMATION_MESSAGE);

//      WtGeneralTools.setJavaLookAndFeel(theme);
    } catch (Exception ex) {
      WtGeneralTools.showErrorMessage(ex);
    }
  }

  public static String encodeUrl(Rule rule) {
    try {
      return URLEncoder.encode(rule.getId(), "utf-8");
    } catch (UnsupportedEncodingException e) {
      throw new RuntimeException(e);
    }
  }

  public static String getExampleSentences(Rule rule, ResourceBundle messages) {
    StringBuilder examples = new StringBuilder(200);
    List<IncorrectExample> incorrectExamples = rule.getIncorrectExamples();
    if (incorrectExamples.size() > 0) {
      String incorrectExample = incorrectExamples.iterator().next().getExample();
      String sentence = incorrectExample.replace("<marker>", "<span style='background-color:#ff8080'>").replace("</marker>", "</span>");
      examples.append("<br/>").append(sentence).append("&nbsp;<span style='color:red;font-style:italic;font-weight:bold'>x</span>");
    }
    List<CorrectExample> correctExamples = rule.getCorrectExamples();
    if (correctExamples.size() > 0) {
      String correctExample = correctExamples.iterator().next().getExample();
      String sentence = correctExample.replace("<marker>", "<span style='background-color:#80ff80'>").replace("</marker>", "</span>");
      examples.append("<br/>").append(sentence).append("&nbsp;<span style='color:green'>✓</span>");
    } else if (incorrectExamples.size() > 0) {
      IncorrectExample incorrectExample = incorrectExamples.iterator().next();
      List<String> corrections = incorrectExample.getCorrections();
      if (!corrections.isEmpty() && !corrections.get(0).isEmpty()) {
        String incorrectSentence = incorrectExamples.iterator().next().getExample();
        String correctedSentence = incorrectSentence.replaceAll("<marker>.*?</marker>",
                "<span style='background-color:#80ff80'>" + corrections.get(0) + "</span>");
        examples.append("<br/>").append(correctedSentence).append("&nbsp;<span style='color:green'>✓</span>");
      }
    }
    if (examples.length() > 0) {
      examples.insert(0, "<br/><br/>" + messages.getString("guiExamples"));
    }
    return examples.toString();
  }

  public static String formatURL(URL url) {
    if (url == null) {
      return "";
    }
    return String.format("<br/><br/><a href=\"%s\">%s</a>",
            url.toExternalForm(), StringUtils.abbreviate(url.toString(), 50));
  }

  public static void setJavaLookAndFeel(int theme) throws Exception {
    switch (theme) {
    case THEME_FLATDARK:
      FlatDarkLaf.setup();
      break;
    case THEME_FLATLIGHT:
      FlatLightLaf.setup();
      break;
    default:
      // System dependent
      // do not set look and feel for on Mac OS X as it causes the following error:
      // soffice[2149:2703] Apple AWT Java VM was loaded on first thread -- can't start AWT.
      if (!System.getProperty("os.name").contains("OS X")) {
         // Cross-Platform Look And Feel @since 3.7
         if (System.getProperty("os.name").contains("Linux")) {
           boolean isGTK = false;
           LookAndFeelInfo[] lookAndFeels = UIManager.getInstalledLookAndFeels();
           if (!(lookAndFeels == null)) {
             for (LookAndFeelInfo lookAndFeel : lookAndFeels) {
               if ("GTK+".equals(lookAndFeel.getName())) {
                 isGTK = true;
                 break;
               }
             }
           }
           if (isGTK) {
             UIManager.setLookAndFeel("com.sun.java.swing.plaf.gtk.GTKLookAndFeel");
           } else {
             UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
           }
         }
         else {
           UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
         }
      } else {
        UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
      }
      break;
    }
  }
}
