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
package org.writingtool.dialogs;

import org.apache.commons.lang3.StringUtils;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;
import org.languagetool.Language;
import org.languagetool.Languages;
import org.languagetool.languagemodel.LuceneLanguageModel;
import org.languagetool.rules.Rule;
import org.languagetool.rules.RuleOption;
import org.writingtool.config.WtCategoryNode;
import org.writingtool.config.WtCheckBoxTreeCellRenderer;
import org.writingtool.config.WtConfiguration;
import org.writingtool.config.WtRuleNode;
import org.writingtool.config.WtSavablePanel;
import org.writingtool.config.WtTreeListener;
import org.writingtool.tools.WtGeneralTools;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;

import javax.swing.*;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.event.TreeModelEvent;
import javax.swing.event.TreeModelListener;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.DefaultTreeModel;
import javax.swing.tree.TreeNode;
import javax.swing.tree.TreePath;
import java.awt.*;
import java.awt.event.*;
import java.io.File;
import java.io.IOException;
import java.text.MessageFormat;
import java.util.*;
import java.util.List;

/**
 * Dialog that offers the available rules so they can be turned on/off
 * individually.
 * 
 * @author Fred Kruse
 */
public class WtConfigurationDialog implements ActionListener {

  private static final String COLOR_LABEL = " \u2588\u2588\u2588 ";  // \u2587 is smaller
  private static final String NO_SELECTED_LANGUAGE = "---";
  private static final String ACTION_COMMAND_OK = "OK";
  private static final String ACTION_COMMAND_CANCEL = "CANCEL";

  private static final int SHIFT1 = 4;
  private static final int SHIFT2 = 20;
  private static final int SHIFT3 = 44;
  private static final int YSHIFT1 = 8;

  private final ResourceBundle messages;
  private final WtConfiguration original;
  private final WtConfiguration config;
  private final Frame owner;
  private final Image ltImage;
  private String dialogTitle;
  private boolean configChanged = false;
  private boolean profileChanged = true;
  private boolean restartShow = false;
  private boolean firstSelection = true;

  private JDialog dialog;
  private JTree[] configTree;
  private DefaultMutableTreeNode[] rootNode;
  private JPanel disabledRulesPanel;
  private JPanel enabledRulesPanel;
  private JTabbedPane tabpane;
  private final List<JPanel> extraPanels = new ArrayList<>();
  private final List<Rule> configurableRules = new ArrayList<>();
  private String category;
  private Rule rule;
  private int tabIndex = 0;

  public WtConfigurationDialog(Frame owner, WtConfiguration config) {
    this(owner, null, null, config);
  }

  public WtConfigurationDialog(Frame owner, Image ltImage, String title, WtConfiguration config) {
    this.owner = owner;
    this.original = config;
    this.config = original.copy(original);
    this.ltImage = ltImage;
    dialogTitle = title;
    messages = WtOfficeTools.getMessageBundle();
  }

  /**
   * Add extra JPanel to this dialog.
   * 
   * If the panel implements {@see SavablePanel}, this dialog will call
   * {@link WtSavablePanel#save} after the user clicks OK.
   * 
   * @param panel the JPanel to be added to this dialog
   * @since 3.4
   */
  void addExtraPanel(JPanel panel) {
    extraPanels.add(panel);
  }

  private DefaultMutableTreeNode createTree(List<Rule> rules, boolean isStyle, String tabName, DefaultMutableTreeNode root) {
    if (root == null) {
      root = new DefaultMutableTreeNode("Rules");
    } else {
      root.removeAllChildren();
    }
    String lastRuleId = null;
    Map<String, DefaultMutableTreeNode> parents = new TreeMap<>();
    for (Rule rule : rules) {
      if((tabName == null && !config.isSpecialTabCategory(rule.getCategory().getName()) &&
          ((isStyle && config.isStyleCategory(rule.getCategory().getName())) ||
         (!isStyle && !config.isStyleCategory(rule.getCategory().getName())))) || 
          (tabName != null && config.isInSpecialTab(rule.getCategory().getName(), tabName))) {
        if (!parents.containsKey(rule.getCategory().getName())) {
          boolean enabled = true;
          if (config.getDisabledCategoryNames() != null && config.getDisabledCategoryNames().contains(rule.getCategory().getName())) {
            enabled = false;
          }
          if(rule.getCategory().isDefaultOff() && (config.getEnabledCategoryNames() == null 
              || !config.getEnabledCategoryNames().contains(rule.getCategory().getName()))) {
            enabled = false;
          }
          DefaultMutableTreeNode categoryNode = new WtCategoryNode(rule.getCategory(), enabled);
          root.add(categoryNode);
          parents.put(rule.getCategory().getName(), categoryNode);
        }
        if (!rule.getId().equals(lastRuleId)) {
          WtRuleNode ruleNode = new WtRuleNode(rule, getEnabledState(rule));
          parents.get(rule.getCategory().getName()).add(ruleNode);
        }
        lastRuleId = rule.getId();
      }
    }
    return root;
  }

  private boolean getEnabledState(Rule rule) {
    boolean ret = true;
    if (config.getDisabledRuleIds().contains(rule.getId())) {
      ret = false;
    }
    if (config.getDisabledCategoryNames().contains(rule.getCategory().getName())) {
      ret = false;
    }
    if ((rule.isDefaultOff() || rule.getCategory().isDefaultOff()) && !config.getEnabledRuleIds().contains(rule.getId())) {
      ret = false;
    }
    if (rule.isOfficeDefaultOff() && !config.getEnabledRuleIds().contains(rule.getId())) {
      ret = false;
    }
    if (rule.isOfficeDefaultOn() && !config.getDisabledRuleIds().contains(rule.getId())) {
      ret = true;
    }
    if (rule.isDefaultOff() && rule.getCategory().isDefaultOff()
            && config.getEnabledRuleIds().contains(rule.getId())) {
      config.getDisabledCategoryNames().remove(rule.getCategory().getName());
    }
    return ret;
  }

  public boolean show(List<Rule> rules) {
    restartShow = false;
    do {
      showPanel(rules);
    } while (restartShow);
    return configChanged;
  }
    
  public void close() {
    dialog.setVisible(false);
  }

  public boolean showPanel(List<Rule> rules) {
    configChanged = false;
    if (original != null && !restartShow) {
      config.restoreState(original);
    }
    restartShow = false;
    dialog = new JDialog(owner, true);
    if (dialogTitle == null) {
      dialogTitle = messages.getString("guiWtConfigWindowTitle");
    }
    dialog.setTitle(dialogTitle);
    if (ltImage != null) {
      ((Frame) dialog.getOwner()).setIconImage(ltImage);
    }
    // close dialog when user presses Escape key:
    KeyStroke stroke = KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0);
    ActionListener actionListener = actionEvent -> dialog.setVisible(false);
    JRootPane rootPane = dialog.getRootPane();
    rootPane.registerKeyboardAction(actionListener, stroke,
        JComponent.WHEN_IN_FOCUSED_WINDOW);

    configurableRules.clear();

    Language lang = config.getLanguage();
    if (lang == null) {
      lang = Languages.getLanguageForLocale(Locale.getDefault());
    }

    String[] specialTabNames = config.getSpecialTabNames();
    int numConfigTrees = 2 + specialTabNames.length;
    configTree = new JTree[numConfigTrees];
    rootNode = new DefaultMutableTreeNode[numConfigTrees];
    JPanel[] checkBoxPanel = new JPanel[numConfigTrees];
    GridBagConstraints cons;

    for (int i = 0; i < numConfigTrees; i++) {
      checkBoxPanel[i] = new JPanel();
      cons = new GridBagConstraints();
      checkBoxPanel[i].setLayout(new GridBagLayout());
      cons.anchor = GridBagConstraints.NORTHWEST;
      cons.gridx = 0;
      cons.weightx = 1.0;
      cons.weighty = 1.0;
      cons.fill = GridBagConstraints.HORIZONTAL;
      Collections.sort(rules, new CategoryComparator());
      if(i == 0) {
        rootNode[i] = createTree(rules, false, null, null);   //  grammar options
      } else if(i == 1) {
        rootNode[i] = createTree(rules, true, null, null);    //  Style options
      } else {
        rootNode[i] = createTree(rules, true, specialTabNames[i - 2], null);    //  Special tab options
      }
      configTree[i] = new JTree(getTreeModel(rootNode[i], rules));
      
      configTree[i].applyComponentOrientation(ComponentOrientation.getOrientation(lang.getLocale()));
  
      configTree[i].setRootVisible(false);
      configTree[i].setEditable(false);
      configTree[i].setCellRenderer(new WtCheckBoxTreeCellRenderer());
      WtTreeListener.install(configTree[i]);
      checkBoxPanel[i].add(configTree[i], cons);
      configTree[i].addMouseListener(getMouseAdapter());
    }
    

    JPanel portPanel = new JPanel();
    portPanel.setLayout(new GridBagLayout());
    cons = new GridBagConstraints();
    cons.insets = new Insets(0, SHIFT1, 0, 0);
    cons.gridx = 0;
    cons.gridy = 0;
    cons.anchor = GridBagConstraints.WEST;
    cons.fill = GridBagConstraints.NONE;
    cons.weightx = 0.0f;
    createOfficeElements(cons, portPanel);

    JPanel subButtonPanel = new JPanel();
    subButtonPanel.setLayout(new GridBagLayout());
    JButton okButton = new JButton(WtGeneralTools.getLabel(messages.getString("guiOKButton")));
    okButton.setMnemonic(WtGeneralTools.getMnemonic(messages.getString("guiOKButton")));
    okButton.setActionCommand(ACTION_COMMAND_OK);
    okButton.addActionListener(this);
    JButton cancelButton = new JButton(WtGeneralTools.getLabel(messages.getString("guiCancelButton")));
    cancelButton.setMnemonic(WtGeneralTools.getMnemonic(messages.getString("guiCancelButton")));
    cancelButton.setActionCommand(ACTION_COMMAND_CANCEL);
    cancelButton.addActionListener(this);
    JButton helpButton = new JButton(WtGeneralTools.getLabel(messages.getString("guiHelpButton")));
    helpButton.setMnemonic(WtGeneralTools.getMnemonic(messages.getString("guiHelpButton")));
    helpButton.setActionCommand("Help");
    helpButton.addActionListener(this);
    cons = new GridBagConstraints();
    subButtonPanel.add(okButton, cons);
    subButtonPanel.add(cancelButton, cons);
    JPanel buttonPanel = new JPanel();
    buttonPanel.setLayout(new GridBagLayout());
    cons = new GridBagConstraints();
    cons.insets = new Insets(0, SHIFT1, 0, 0);
    cons.gridx = 0;
    cons.gridy = 0;
    cons.anchor = GridBagConstraints.WEST;
//    cons.fill = GridBagConstraints.HORIZONTAL;
    cons.weightx = 1.0f;
    cons.weighty = 0.0f;
    buttonPanel.add(helpButton, cons);
    cons.gridx++;
    cons.anchor = GridBagConstraints.EAST;
    buttonPanel.add(subButtonPanel, cons);

    tabpane = new JTabbedPane();

//  Profile tab    
    JPanel jProfilePane = new JPanel();
    jProfilePane.setLayout(new GridBagLayout());
    cons = new GridBagConstraints();
    cons.insets = new Insets(4, 4, 4, 4);

    cons.gridx = 0;
    cons.gridy = 0;
    cons.weightx = 10.0f;
    cons.weighty = 4.0f;
    cons.fill = GridBagConstraints.BOTH;
    cons.anchor = GridBagConstraints.NORTHWEST;
    
    jProfilePane.add(new JScrollPane(getProfilePanel(rules)), cons);
    
    //  Disabled default rules
    cons.fill = GridBagConstraints.BOTH;
    cons.weighty = 0.0f;
    cons.gridy++;
    cons.insets = new Insets(16, 4, 0, 8);
    jProfilePane.add(new JLabel(addColonToMessageString("guiDisabledDefaultRules")), cons);
    cons.insets = new Insets(8, 4, 0, 8);
    cons.gridy++;
    cons.weighty = 3.0f;
    disabledRulesPanel = getChangedRulesPanel(rules, false, null);
    jProfilePane.add(new JScrollPane(disabledRulesPanel), cons);
    
    //  Enabled optional rules
    cons.gridy++;
    cons.insets = new Insets(16, 4, 0, 8);
    cons.weighty = 0.0f;
    jProfilePane.add(new JLabel(addColonToMessageString("guiEnabledOptionalRules")), cons);
    cons.insets = new Insets(8, 4, 0, 8);
    cons.gridy++;
    cons.weighty = 3.0f;
    enabledRulesPanel = getChangedRulesPanel(rules, true, null);
    jProfilePane.add(new JScrollPane(enabledRulesPanel), cons);
    jProfilePane.setName(messages.getString("guiProfiles"));
    
//  General tab
    JPanel jPane = new JPanel();
    jPane.setLayout(new GridBagLayout());
    cons = new GridBagConstraints();
    cons.insets = new Insets(4, 4, 4, 4);

    cons.gridx = 0;
    cons.gridy = 0;
    cons.weightx = 10.0f;
    cons.weighty = 0.0f;
    cons.fill = GridBagConstraints.NONE;
    cons.anchor = GridBagConstraints.NORTHWEST;

    cons.gridy++;
    cons.anchor = GridBagConstraints.WEST;
    jPane.add(portPanel, cons);
    cons.fill = GridBagConstraints.HORIZONTAL;
    cons.anchor = GridBagConstraints.WEST;
    for(JPanel extra : extraPanels) {
      //in case it wasn't in a containment hierarchy when user changed L&F
      SwingUtilities.updateComponentTreeUI(extra);
      cons.gridy++;
      jPane.add(extra, cons);
    }

    cons.gridy++;
    cons.fill = GridBagConstraints.BOTH;
    cons.weighty = 1.0f;
    jPane.add(new JPanel(), cons);
    
    tabpane.addTab(messages.getString("guiGeneral"), new JScrollPane(jPane));

//  Grammar rules tab    
    jPane = new JPanel();
    jPane.setLayout(new GridBagLayout());
    cons = new GridBagConstraints();
    cons.insets = new Insets(4, 4, 4, 4);
    cons.gridx = 0;
    cons.gridy = 0;
    cons.weightx = 10.0f;
    cons.weighty = 10.0f;
    cons.fill = GridBagConstraints.BOTH;
    jPane.add(new JScrollPane(checkBoxPanel[0]), cons);
    cons.weightx = 0.0f;
    cons.weighty = 0.0f;

    cons.gridx = 0;
    cons.gridy++;
    cons.fill = GridBagConstraints.NONE;
    cons.anchor = GridBagConstraints.LINE_END;
    jPane.add(getTreeButtonPanel(0), cons);
    cons.fill = GridBagConstraints.HORIZONTAL;
    cons.anchor = GridBagConstraints.WEST;
    cons.gridx = 0;
    cons.gridy++;
    jPane.add(getRuleOptionsPanel(0), cons);

    tabpane.addTab(messages.getString("guiGrammarRules"), jPane);
    
//  Style rules tab    
    jPane = new JPanel();
    jPane.setLayout(new GridBagLayout());
    cons = new GridBagConstraints();
    cons.insets = new Insets(4, 4, 4, 4);
    cons.gridx = 0;
    cons.gridy = 0;
    cons.weightx = 10.0f;
    cons.weighty = 10.0f;
    cons.fill = GridBagConstraints.BOTH;
    jPane.add(new JScrollPane(checkBoxPanel[1]), cons);
    cons.weightx = 0.0f;
    cons.weighty = 0.0f;

    cons.gridx = 0;
    cons.gridy++;
    cons.fill = GridBagConstraints.NONE;
    cons.anchor = GridBagConstraints.LINE_END;
    jPane.add(getTreeButtonPanel(1), cons);
    cons.fill = GridBagConstraints.HORIZONTAL;
    cons.anchor = GridBagConstraints.WEST;
    cons.gridx = 0;
    cons.gridy++;
    jPane.add(getRuleOptionsPanel(1), cons);

    tabpane.addTab(messages.getString("guiStyleRules"), jPane);

//  Style special tabs (optional)
    for (int i = 0; i < specialTabNames.length; i++) {
      jPane = new JPanel();
      jPane.setLayout(new GridBagLayout());
      cons = new GridBagConstraints();
      cons.insets = new Insets(4, 4, 4, 4);
      cons.gridx = 0;
      cons.gridy = 0;
      cons.weightx = 10.0f;
      cons.weighty = 10.0f;
      cons.fill = GridBagConstraints.BOTH;
      jPane.add(new JScrollPane(checkBoxPanel[i + 2]), cons);
      cons.weightx = 0.0f;
      cons.weighty = 0.0f;
  
      cons.gridx = 0;
      cons.gridy++;
      cons.fill = GridBagConstraints.NONE;
      cons.anchor = GridBagConstraints.LINE_END;
      jPane.add(getTreeButtonPanel(i + 2), cons);
  
      cons.fill = GridBagConstraints.HORIZONTAL;
      cons.anchor = GridBagConstraints.WEST;
      cons.gridx = 0;
      cons.gridy++;
      jPane.add(getRuleOptionsPanel(i + 2), cons);

      tabpane.addTab(specialTabNames[i], jPane);
    }
    
    //    Default color options tab
    if (!config.onlySingleParagraphMode()) {
      //  NOTE: onlySingleParagraphMode is used for OO and old LO installation
      //        different colors are not supported by such applications
      jPane = new JPanel();
      jPane.setLayout(new GridBagLayout());
      cons = new GridBagConstraints();
      cons.insets = new Insets(4, 4, 4, 4);
      cons.gridx = 0;
      cons.gridy = 0;
      cons.weightx = 10.0f;
      cons.weighty = 10.0f;
      cons.anchor = GridBagConstraints.NORTHWEST;
      cons.fill = GridBagConstraints.NONE;
      jPane.add(getOfficeDefaultColorPanel(), cons);
      String label = messages.getString("guiColorTabLabel");
      tabpane.add(label, new JScrollPane(jPane));
    }
    
    //    technical options tab
    jPane = new JPanel();
    jPane.setLayout(new GridBagLayout());
    cons = new GridBagConstraints();
    cons.insets = new Insets(4, 4, 4, 4);
    cons.gridx = 0;
    cons.gridy = 0;
    cons.weightx = 10.0f;
    cons.weighty = 10.0f;
    cons.anchor = GridBagConstraints.NORTHWEST;
    cons.fill = GridBagConstraints.NONE;
    jPane.add(getOfficeTechnicalElements(), cons);
    String label = messages.getString("guiTechnicalSettings");
    if (label.endsWith(":")) {
      label = label.substring(0, label.length() - 1);
    }
    tabpane.add(label, new JScrollPane(jPane));
    
    //    AI options tab
    //    onlySingleParagraphMode doesn't support cache so AI can't be used
    if (!config.onlySingleParagraphMode()) {
      jPane = new JPanel();
      jPane.setLayout(new GridBagLayout());
      jPane.add(getOfficeAiElements(), cons);
      label = messages.getString("guiAiSupportSettings");
      tabpane.add(label, new JScrollPane(jPane));
    }

    Container contentPane = dialog.getContentPane();
    contentPane.setLayout(new GridBagLayout());
    cons = new GridBagConstraints();
    cons.insets = new Insets(4, 4, 4, 4);
    cons.gridx = 0;
    cons.gridy = 0;
    cons.weightx = 10.0f;
    cons.weighty = 10.0f;
    cons.fill = GridBagConstraints.BOTH;
    cons.anchor = GridBagConstraints.NORTHWEST;
    contentPane.add(tabpane, cons);
    cons.weightx = 1.0f;
    cons.weighty = 0.0f;
    cons.gridy++;
//    cons.fill = GridBagConstraints.NONE;
//    cons.anchor = GridBagConstraints.EAST;
    contentPane.add(buttonPanel, cons);

    dialog.pack();
    // center on screen:
    Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
    Dimension frameSize = dialog.getSize();
    dialog.setLocation(screenSize.width / 2 - frameSize.width / 2,
        screenSize.height / 2 - frameSize.height / 2);
    dialog.setLocationByPlatform(true);
    //  add Profile tab after dimension was set
    tabpane.add(jProfilePane, 0);
    tabpane.setSelectedIndex(tabIndex);
    for(JPanel extra : this.extraPanels) {
      if(extra instanceof WtSavablePanel) {
        ((WtSavablePanel) extra).componentShowing();
      }
    }
    dialog.setAutoRequestFocus(true);
    dialog.setAlwaysOnTop(true);
    dialog.setVisible(true);
    dialog.toFront();
    return configChanged;
  }

  private JPanel getOfficeDefaultColorPanel() {
    // default color settings
    JPanel panel = new JPanel();
    panel.setLayout(new GridBagLayout());
    GridBagConstraints cons = new GridBagConstraints();
    cons.insets = new Insets(0, SHIFT1, 0, 0);
    cons.gridx = 0;
    cons.gridy = 0;
    cons.anchor = GridBagConstraints.WEST;
    cons.fill = GridBagConstraints.NONE;
    cons.weightx = 0.0f;
    
    JPanel themePanel = new JPanel();
    themePanel.setLayout(new GridBagLayout());
    GridBagConstraints cons3 = new GridBagConstraints();
    cons3.insets = new Insets(0, 0, 0, 0);
    cons3.gridx = 0;
    cons3.gridy = 0;
    cons3.anchor = GridBagConstraints.WEST;
    cons3.fill = GridBagConstraints.NONE;
    cons3.weightx = 0.0f;
    themePanel.add(new JLabel(messages.getString("guiThemeLabel") + ": "), cons3);
    String[] themes = { "System", "FlatDark", "FlatLight" };
    JComboBox<String> themeBox = new JComboBox<>(themes);
    themeBox.setSelectedIndex(config.getThemeSelection());
    themeBox.addItemListener(e -> {
      if (e.getStateChange() == ItemEvent.SELECTED) {
        if (themeBox.getSelectedIndex() != config.getThemeSelection()) {
          config.setThemeSelection(themeBox.getSelectedIndex());
          try {
            WtGeneralTools.setJavaLookAndFeel(themeBox.getSelectedIndex());
          } catch (Exception e1) {
            WtMessageHandler.showError(e1);
          }
          tabIndex = 4 + config.getSpecialTabNames().length;
          restartShow = true;
          dialog.setVisible(false);
        }
      }
    });
    cons3.gridx++;
    themePanel.add(themeBox, cons3);

    JPanel customPanel = new JPanel();
    customPanel.setLayout(new GridBagLayout());
    GridBagConstraints cons1 = new GridBagConstraints();
    cons1.gridx = 0;
    cons1.gridy = 0;
    cons1.weightx = 0.0f;
    cons1.fill = GridBagConstraints.NONE;
    cons1.anchor = GridBagConstraints.NORTHWEST;
    
    JLabel[] jLabels = new JLabel[3];
    JLabel[] underlineLabels = new JLabel[3];
    JButton[] changeButtons  = new JButton[3];
    
    List<Color> defaultColors = config.getUnderlineDefaultColors();
    
    for (int i = 0; i < 3; i++) {
      cons1.gridx = 0;
      cons1.gridy++;
      cons1.insets = new Insets(3, 0, 0, 0);
      if (i == 0) {
        jLabels[i] = new JLabel(messages.getString("guiCustomColorGrammar") + ": ");
      } else if (i == 1) {
        jLabels[i] = new JLabel(messages.getString("guiCustomColorStyle") + ": ");
      } else {
        jLabels[i] = new JLabel(messages.getString("guiCustomColorOptional") + ": ");
      }
      jLabels[i].setVerticalAlignment(SwingConstants.CENTER);
      customPanel.add(jLabels[i], cons1);
      
      underlineLabels[i] = new JLabel(COLOR_LABEL);
      underlineLabels[i].setVerticalAlignment(SwingConstants.CENTER);
      int n = i;
      cons1.gridx++;
      customPanel.add(underlineLabels[i], cons1);
  
      changeButtons[i] = new JButton(messages.getString("guiUColorChange"));
      changeButtons[i].setVerticalAlignment(SwingConstants.CENTER);
      changeButtons[i].addActionListener(e -> {
        try {
//          int theme = WtDocumentsHandler.getJavaLookAndFeelSet();
//          WtGeneralTools.setJavaLookAndFeel(WtGeneralTools.THEME_SYSTEM);
          Color oldColor = underlineLabels[n].getForeground();
          Color newColor = JColorChooser.showDialog(dialog, messages.getString("guiUColorDialogHeader"), oldColor);
          if(newColor != null && !newColor.equals(oldColor)) {
            underlineLabels[n].setForeground(newColor);
            defaultColors.set(n, newColor);
            config.setUnderlineDefaultColor(defaultColors);
          }
//          WtGeneralTools.setJavaLookAndFeel(theme);
        } catch (Exception e1) {
          WtMessageHandler.printException(e1);
        }
      });
      cons1.gridx++;
      cons1.insets = new Insets(1, 0, 0, 0);
      customPanel.add(changeButtons[i], cons1);
    
      underlineLabels[i].setForeground(defaultColors.get(i));
      underlineLabels[i].setBackground(defaultColors.get(i));
    }

    JPanel radioPanel = new JPanel();
    radioPanel.setLayout(new GridBagLayout());
    GridBagConstraints cons2 = new GridBagConstraints();
    cons2.insets = new Insets(0, SHIFT1, 0, 0);
    cons2.gridx = 0;
    cons2.gridy = 0;
    cons2.anchor = GridBagConstraints.WEST;
    cons2.fill = GridBagConstraints.NONE;
    cons2.weightx = 0.0f;
    
    JRadioButton[] radioButtons = new JRadioButton[5];
    ButtonGroup colorSelectionGroup = new ButtonGroup();
    radioButtons[0] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiWTColorPalette")));
    radioButtons[0].addActionListener(e -> {
      config.setColorSelection(0);
      for (JButton button : changeButtons) {
        button.setEnabled(false);
      }
    });
    radioButtons[1] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiBlueColorPalette")));
    radioButtons[1].addActionListener(e -> {
      config.setColorSelection(1);
      for (JButton button : changeButtons) {
        button.setEnabled(false);
      }
    });
    radioButtons[2] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiLTColorPalette")));
    radioButtons[2].addActionListener(e -> {
      config.setColorSelection(2);
      for (JButton button : changeButtons) {
        button.setEnabled(false);
      }
    });
    radioButtons[3] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiDarkColorPalette")));
    radioButtons[3].addActionListener(e -> {
      config.setColorSelection(3);
      for (JButton button : changeButtons) {
        button.setEnabled(false);
      }
    });
    radioButtons[4] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiCustomColorPalete") + ":"));
    radioButtons[4].addActionListener(e -> {
      config.setColorSelection(99);
      for (JButton button : changeButtons) {
        button.setEnabled(true);
      }
    });

    for (int i = 0; i < 5; i++) {
      if ((i == radioButtons.length - 1 && config.getColorSelection() == 99) || config.getColorSelection() == i) {
        radioButtons[i].setSelected(true);
      } else {
        radioButtons[i].setSelected(false);
      }
      colorSelectionGroup.add(radioButtons[i]);
    }
    
    for (JButton button : changeButtons) {
      button.setEnabled(config.getColorSelection() == 99);
    }
    
    cons.insets = new Insets(4, SHIFT1, 0, 0);
    for (int i = 0; i < 5; i++) {
      cons2.gridy++;
      radioPanel.add(radioButtons[i], cons2);
    }
    panel.add(themePanel, cons);
    cons.gridy++;
    panel.add(new JLabel(" "), cons);
    cons.gridy++;
    panel.add(new JLabel(messages.getString("guiColorSelectionLabel") + ":"), cons);
    cons.gridy++;
    panel.add(radioPanel, cons);
    cons.insets = new Insets(4, SHIFT3, 0, 0);
    cons.gridy++;
    panel.add(customPanel, cons);
    customPanel.setEnabled(config.getColorSelection() == 99);
    return panel;
  }

  private void addOfficeLanguageElements(GridBagConstraints cons, JPanel portPanel) {
    JPanel languagePanel = new JPanel();
    languagePanel.setLayout(new GridBagLayout());
    GridBagConstraints cons1 = new GridBagConstraints();
    cons1.insets = new Insets(0, 0, 0, 0);
    cons1.gridx = 0;
    cons1.gridy = 0;
    cons1.anchor = GridBagConstraints.WEST;
    cons1.fill = GridBagConstraints.NONE;
    cons1.weightx = 0.0f;
    JRadioButton[] radioButtons = new JRadioButton[2];
    ButtonGroup numParaGroup = new ButtonGroup();
    radioButtons[0] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiUseDocumentLanguage")));
    radioButtons[0].setActionCommand("DocLang");
    radioButtons[0].setSelected(true);

    radioButtons[1] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiSetLanguageTo")));
    radioButtons[1].setActionCommand("SelectLang");

    JComboBox<String> fixedLanguageBox = new JComboBox<>(getPossibleLanguages(false));
    if (config.getFixedLanguage() != null) {
      fixedLanguageBox.setSelectedItem(config.getFixedLanguage().getTranslatedName(messages));
    }
    fixedLanguageBox.addItemListener(e -> {
      if (e.getStateChange() == ItemEvent.SELECTED) {
        Language fixedLanguage;
        if (fixedLanguageBox.getSelectedItem() instanceof String) {
          fixedLanguage = getLanguageForLocalizedName(fixedLanguageBox.getSelectedItem().toString());
        } else {
          fixedLanguage = (Language) fixedLanguageBox.getSelectedItem();
        }
        config.setFixedLanguage(fixedLanguage);
        config.setUseDocLanguage(false);
        radioButtons[1].setSelected(true);
      }
    });
    
    for (int i = 0; i < 2; i++) {
      numParaGroup.add(radioButtons[i]);
    }
    
    if (config.getUseDocLanguage()) {
      radioButtons[0].setSelected(true);
    } else {
      radioButtons[1].setSelected(true);
    }

    radioButtons[0].addActionListener(e -> config.setUseDocLanguage(true));
    
    radioButtons[1].addActionListener(e -> {
      config.setUseDocLanguage(false);
      Language fixedLanguage;
      if (fixedLanguageBox.getSelectedItem() instanceof String) {
        fixedLanguage = getLanguageForLocalizedName(fixedLanguageBox.getSelectedItem().toString());
      } else {
        fixedLanguage = (Language) fixedLanguageBox.getSelectedItem();
      }
      config.setFixedLanguage(fixedLanguage);
    });
    languagePanel.add(radioButtons[0], cons1);
    cons1.gridy++;
    languagePanel.add(radioButtons[1], cons1);
    cons1.gridx = 1;
    languagePanel.add(fixedLanguageBox, cons1);

    cons.insets = new Insets(0, SHIFT1, 0, 0);
    cons.gridx = 0;
    cons.gridy++;
    portPanel.add(languagePanel, cons);
  }
/*
  private void addOfficeTextruleElements(GridBagConstraints cons, JPanel portPanel) {
    int numParaCheck = config.getNumParasToCheck();
    boolean useTextLevelQueue = config.useTextLevelQueue();
    JRadioButton[] radioButtons = new JRadioButton[3];
    ButtonGroup numParaGroup = new ButtonGroup();
    radioButtons[0] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiTextCheckMode")));
    radioButtons[0].setActionCommand("FullTextCheck");
    
    radioButtons[1] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiParagraphCheckMode")));
    radioButtons[1].setActionCommand("ParagraphCheck");

    radioButtons[2] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiDeveloperModeCheck")));
    radioButtons[2].setActionCommand("NParagraphCheck");

    JTextField numParaField = new JTextField(Integer.toString(5), 2);
    numParaField.setEnabled(radioButtons[2].isSelected());
    numParaField.setMinimumSize(new Dimension(30, 25));
    
    for (int i = 0; i < 3; i++) {
      numParaGroup.add(radioButtons[i]);
    }

     // NOTE: The flatparagraph iterator doesn't work for OpenOffice (OO).
     //       So no support of cache is possible for OO
    if (numParaCheck == 0 || config.onlySingleParagraphMode()) {
      radioButtons[1].setSelected(true);
      numParaField.setEnabled(false);
      config.setUseTextLevelQueue(false);
      if (config.onlySingleParagraphMode()) {
        radioButtons[0].setEnabled(false);
        radioButtons[2].setEnabled(false);
      }
    } else if (useTextLevelQueue) {
      radioButtons[0].setSelected(true);
      numParaField.setEnabled(false);
      config.setNumParasToCheck(-2);
    } else {
      radioButtons[2].setSelected(true);
      numParaField.setText(Integer.toString(numParaCheck));
      numParaField.setEnabled(true);
    }

    radioButtons[0].addActionListener(e -> {
      numParaField.setEnabled(false);
      config.setNumParasToCheck(-2);
      config.setUseTextLevelQueue(true);
    });
    
    radioButtons[1].addActionListener(e -> {
      numParaField.setEnabled(false);
      config.setNumParasToCheck(0);
      config.setUseTextLevelQueue(false);
    });
    
    radioButtons[2].addActionListener(e -> {
      int numParaCheck1 = Integer.parseInt(numParaField.getText());
      if (numParaCheck1 < -2) numParaCheck1 = -2;
      else if (numParaCheck1 > 99) numParaCheck1 = 99;
      config.setNumParasToCheck(numParaCheck1);
      numParaField.setForeground(Color.BLACK);
      numParaField.setText(Integer.toString(numParaCheck1));
      numParaField.setEnabled(true);
      config.setUseTextLevelQueue(false);
    });
    
    numParaField.getDocument().addDocumentListener(new DocumentListener() {
      @Override
      public void insertUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void removeUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void changedUpdate(DocumentEvent e) {
        try {
          int numParaCheck = Integer.parseInt(numParaField.getText());
          if (numParaCheck > -3 && numParaCheck < 99) {
            numParaField.setForeground(Color.BLACK);
            config.setNumParasToCheck(numParaCheck);
          } else {
            numParaField.setForeground(Color.RED);
          }
        } catch (NumberFormatException ex) {
          numParaField.setForeground(Color.RED);
        }
      }
    });

    JLabel textChangedLabel = new JLabel(WtGeneralTools.getLabel(messages.getString("guiSentenceExceedingRules")));
    cons.gridy++;
    portPanel.add(textChangedLabel, cons);
    
    JPanel radioPanel = new JPanel();
    radioPanel.setLayout(new GridBagLayout());
    GridBagConstraints cons1 = new GridBagConstraints();
    cons1.insets = new Insets(0, 0, 0, 0);
    cons1.gridx = 0;
    cons1.gridy = 0;
    cons1.anchor = GridBagConstraints.WEST;
    cons1.fill = GridBagConstraints.NONE;
    cons1.weightx = 0.0f;
    for (int i = 0; i < 3; i++) {
      radioPanel.add(radioButtons[i], cons1);
      if (i < 2) cons1.gridy++;
    }
    cons1.gridx = 1;
    radioPanel.add(numParaField, cons1);
    cons.insets = new Insets(0, SHIFT2, 0, 0);
    cons.gridy++;
    portPanel.add(radioPanel, cons);
  }
*/  
  private JPanel getOfficeTechnicalElements() {
    // technical settings
    JPanel panel = new JPanel();
    panel.setLayout(new GridBagLayout());
    GridBagConstraints cons = new GridBagConstraints();
    cons.insets = new Insets(0, SHIFT1, 0, 0);
    cons.gridx = 0;
    cons.gridy = 0;
    cons.anchor = GridBagConstraints.WEST;
    cons.fill = GridBagConstraints.NONE;
    cons.weightx = 0.0f;
    JPanel ngramPanel = new JPanel();
    ngramPanel.setLayout(new GridBagLayout());
    GridBagConstraints cons11 = new GridBagConstraints();
    cons11.insets = new Insets(0, 0, 0, 0);
    cons11.gridx = 0;
    cons11.gridy = 0;
    cons11.anchor = GridBagConstraints.WEST;
    cons11.fill = GridBagConstraints.NONE;
    cons11.weightx = 0.0f;
    addNgramPanel(cons11, ngramPanel);
    JCheckBox paragraphModeBox = new JCheckBox(WtGeneralTools.getLabel(messages.getString("guiParagraphCheckMode")));
    JCheckBox saveCacheBox = new JCheckBox(WtGeneralTools.getLabel(messages.getString("guiSaveCacheToFile")));
    JTextField otherServerNameField = new JTextField(config.getServerUrl() ==  null ? "" : config.getServerUrl(), 25);
    otherServerNameField.setMinimumSize(new Dimension(100, 25));
    otherServerNameField.getDocument().addDocumentListener(new DocumentListener() {
      @Override
      public void insertUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void removeUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void changedUpdate(DocumentEvent e) {
        String serverName = otherServerNameField.getText();
        serverName = serverName.trim();
        if(serverName.isEmpty()) {
          serverName = null;
        }
        if (config.isValidServerUrl(serverName)) {
          otherServerNameField.setForeground(Color.BLACK);
          config.setOtherServerUrl(serverName);
        } else {
          otherServerNameField.setForeground(Color.RED);
        }
      }
    });
/*
    JCheckBox useServerBox = new JCheckBox(WtGeneralTools.getLabel(messages.getString("guiUseServer")) + " ");
    useServerBox.setSelected(config.useOtherServer());
    useServerBox.addItemListener(e -> {
      int select = WtOptionPane.OK_OPTION;
      boolean selected = useServerBox.isSelected();
      if(selected && firstSelection) {
        select = showRemoteServerHint(useServerBox, true);
        firstSelection = false;
      } else {
        firstSelection = true;
      }
      if(select == WtOptionPane.OK_OPTION) {
        useServerBox.setSelected(selected);
        config.setUseOtherServer(useServerBox.isSelected());
        otherServerNameField.setEnabled(useServerBox.isSelected());
      } else {
        useServerBox.setSelected(false);
        firstSelection = true;
      }
    });
*/
    JLabel usernameLabel = new JLabel(WtGeneralTools.getLabel(messages.getString("guiPremiumUsername")));

    JTextField usernameField = new JTextField(config.getRemoteUsername() ==  null ? "" : config.getRemoteUsername(), 25);
    usernameField.setMinimumSize(new Dimension(100, 25));
    usernameField.getDocument().addDocumentListener(new DocumentListener() {
      @Override
      public void insertUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void removeUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void changedUpdate(DocumentEvent e) {
        String username = usernameField.getText();
        username = username.trim();
        if(username.isEmpty()) {
          username = null;
        }
        if (username != null) {
          config.setRemoteUsername(username);
        }
      }
    });

    JLabel apiKeyLabel = new JLabel(WtGeneralTools.getLabel(messages.getString("guiPremiumApiKey")));

    JTextField apiKeyField = new JTextField(config.getRemoteApiKey() ==  null ? "" : config.getRemoteApiKey(), 25);
    apiKeyField.setMinimumSize(new Dimension(100, 25));
    apiKeyField.getDocument().addDocumentListener(new DocumentListener() {
      @Override
      public void insertUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void removeUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void changedUpdate(DocumentEvent e) {
        String apiKey = apiKeyField.getText();
        apiKey = apiKey.trim();
        if(apiKey.isEmpty()) {
          apiKey = null;
        }
        if (apiKey != null) {
          config.setRemoteApiKey(apiKey);
        }
      }
    });
/*
    JCheckBox isPremiumBox = new JCheckBox(WtGeneralTools.getLabel(messages.getString("guiUsePremiumAccount")) + " ");
    isPremiumBox.setSelected(config.isPremium());
    isPremiumBox.addItemListener(e -> {
      boolean selected = isPremiumBox.isSelected();
      config.setPremium(selected);
      usernameLabel.setEnabled(selected);
      usernameField.setEnabled(selected);
      apiKeyLabel.setEnabled(selected);
      apiKeyField.setEnabled(selected);
    });
*/    

    if (config.getNumParasToCheck() == 0 || config.onlySingleParagraphMode()) {
      config.setMultiThreadLO(false);
      config.setUseTextLevelQueue(false);
      config.setNumParasToCheck(0);
    }
    
    JRadioButton[] localeRemoteCheckButtons = new JRadioButton[2];
    ButtonGroup localeRemoteCheckGroup = new ButtonGroup();

    JRadioButton[] localeCheckButtons = new JRadioButton[3];
    ButtonGroup localeCheckGroup = new ButtonGroup();
    localeCheckButtons[0] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiOneThread")));
    localeCheckButtons[0].addActionListener(e -> {
      otherServerNameField.setEnabled(false);
      usernameLabel.setEnabled(false);
      usernameField.setEnabled(false);
      apiKeyLabel.setEnabled(false);
      apiKeyField.setEnabled(false);
      config.setMultiThreadLO(false);
      config.setUseTextLevelQueue(false);
      config.setRemoteCheck(false);
    });
    localeCheckButtons[1] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiTwoThreads")));
    localeCheckButtons[1].addActionListener(e -> {
      otherServerNameField.setEnabled(false);
      usernameLabel.setEnabled(false);
      usernameField.setEnabled(false);
      apiKeyLabel.setEnabled(false);
      apiKeyField.setEnabled(false);
      config.setMultiThreadLO(false);
      config.setUseTextLevelQueue(true);
      config.setRemoteCheck(false);
    });
    localeCheckButtons[2] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiIsMultiThread")));
    localeCheckButtons[2].addActionListener(e -> {
      otherServerNameField.setEnabled(false);
      usernameLabel.setEnabled(false);
      usernameField.setEnabled(false);
      apiKeyLabel.setEnabled(false);
      apiKeyField.setEnabled(false);
      config.setMultiThreadLO(true);
      config.setUseTextLevelQueue(true);
      config.setRemoteCheck(false);
    });
    JRadioButton[] remoteCheckButtons = new JRadioButton[3];
    ButtonGroup remoteCheckGroup = new ButtonGroup();
    remoteCheckButtons[0] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiUseLtServer")));
    remoteCheckButtons[0].addActionListener(e -> {
      otherServerNameField.setEnabled(false);
      usernameLabel.setEnabled(false);
      usernameField.setEnabled(false);
      apiKeyLabel.setEnabled(false);
      apiKeyField.setEnabled(false);
      config.setMultiThreadLO(false);
      config.setRemoteCheck(true);
      config.setPremium(false);
      config.setUseOtherServer(false);
    });
    remoteCheckButtons[1] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiUsePremiumAccount")));
    remoteCheckButtons[1].addActionListener(e -> {
      otherServerNameField.setEnabled(false);
      usernameLabel.setEnabled(true);
      usernameField.setEnabled(true);
      apiKeyLabel.setEnabled(true);
      apiKeyField.setEnabled(true);
      config.setMultiThreadLO(false);
      config.setRemoteCheck(true);
      config.setPremium(true);
      config.setUseOtherServer(false);
    });
    remoteCheckButtons[2] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiUseOtherServer")));
    remoteCheckButtons[2].addActionListener(e -> {
      int select = WtOptionPane.OK_OPTION;
      if(firstSelection) {
        select = showRemoteServerHint(remoteCheckButtons[2], true);
        firstSelection = false;
      } else {
        firstSelection = true;
      }
      if(select == WtOptionPane.OK_OPTION) {
        config.setUseOtherServer(true);
        otherServerNameField.setEnabled(true);
        usernameLabel.setEnabled(false);
        usernameField.setEnabled(false);
        apiKeyLabel.setEnabled(false);
        apiKeyField.setEnabled(false);
        config.setMultiThreadLO(false);
        config.setRemoteCheck(true);
        config.setPremium(false);
        config.setUseOtherServer(true);
      } else {
        localeRemoteCheckButtons[0].setSelected(true);
        firstSelection = true;
      }
    });
    
    localeRemoteCheckButtons[0] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiUseLocaleComputer")));
    localeRemoteCheckButtons[0].addActionListener(e -> {
      otherServerNameField.setEnabled(false);
      usernameLabel.setEnabled(false);
      usernameField.setEnabled(false);
      apiKeyLabel.setEnabled(false);
      apiKeyField.setEnabled(false);
      config.setMultiThreadLO(localeCheckButtons[2].isSelected());
      config.setUseTextLevelQueue(!localeCheckButtons[0].isSelected());
      config.setRemoteCheck(false);
      if (config.getNumParasToCheck() == 0) {
        localeCheckButtons[0].setEnabled(true);
      } else {
        for (int i = 0; i < 3; i++) {
          localeCheckButtons[i].setEnabled(true);
        }
      }
      for (int i = 0; i < 3; i++) {
        remoteCheckButtons[i].setEnabled(false);
      }
    });
    localeRemoteCheckButtons[1] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiUseRemoteServer")));
    localeRemoteCheckButtons[1].addActionListener(e -> {
      int select = WtOptionPane.OK_OPTION;
      boolean selected = localeRemoteCheckButtons[1].isSelected();
      if(selected && firstSelection) {
        select = showRemoteServerHint(localeRemoteCheckButtons[1], false);
        firstSelection = false;
      } else {
        firstSelection = true;
      }
      if(select == WtOptionPane.OK_OPTION) {
//        typeOfCheckButtons[2].setSelected(selected);
        otherServerNameField.setEnabled(config.useOtherServer());
        usernameLabel.setEnabled(config.isPremium());
        usernameField.setEnabled(config.isPremium());
        apiKeyLabel.setEnabled(config.isPremium());
        apiKeyField.setEnabled(config.isPremium());
        config.setMultiThreadLO(false);
        config.setRemoteCheck(true);
        for (int i = 0; i < 3; i++) {
          localeCheckButtons[i].setEnabled(false);
        }
        for (int i = 0; i < 3; i++) {
          remoteCheckButtons[i].setEnabled(true);
        }
      } else {
        localeRemoteCheckButtons[0].setSelected(true);
        firstSelection = true;
      }
    });
    for (int i = 0; i < 2; i++) {
      localeRemoteCheckGroup.add(localeRemoteCheckButtons[i]);
    }
    for (int i = 0; i < 3; i++) {
      localeCheckGroup.add(localeCheckButtons[i]);
    }
    for (int i = 0; i < 3; i++) {
      remoteCheckGroup.add(remoteCheckButtons[i]);
    }
    if (config.useOtherServer()) {
      remoteCheckButtons[2].setSelected(true);
    } else if (config.isPremium()) {
      remoteCheckButtons[1].setSelected(true);
    } else {
      remoteCheckButtons[0].setSelected(true);
    }
    if (config.useTextLevelQueue()) {
      if (config.isMultiThread()) {
        localeCheckButtons[2].setSelected(true);
      } else {
        localeCheckButtons[1].setSelected(true);
      }
    } else {
      localeCheckButtons[0].setSelected(true);
      config.setMultiThreadLO(false);
    }
    if (config.doRemoteCheck()) {
      localeRemoteCheckButtons[1].setSelected(true);
      otherServerNameField.setEnabled(config.useOtherServer());
      usernameLabel.setEnabled(config.isPremium());
      usernameField.setEnabled(config.isPremium());
      apiKeyLabel.setEnabled(config.isPremium());
      apiKeyField.setEnabled(config.isPremium());
      config.setMultiThreadLO(false);
      for (int i = 0; i < 3; i++) {
        localeCheckButtons[i].setEnabled(false);
      }
      for (int i = 0; i < 3; i++) {
        remoteCheckButtons[i].setEnabled(true);
      }
    } else {
      localeRemoteCheckButtons[0].setSelected(true);
      otherServerNameField.setEnabled(false);
      usernameLabel.setEnabled(false);
      usernameField.setEnabled(false);
      apiKeyLabel.setEnabled(false);
      apiKeyField.setEnabled(false);
      ngramPanel.setVisible(true);
      config.setRemoteCheck(false);
      if (config.getNumParasToCheck() == 0) {
        localeCheckButtons[0].setEnabled(true);
        for (int i = 1; i < 3; i++) {
          localeCheckButtons[i].setEnabled(false);
        }
      } else {
        for (int i = 0; i < 3; i++) {
          localeCheckButtons[i].setEnabled(true);
        }
      }
      for (int i = 0; i < 3; i++) {
        remoteCheckButtons[i].setEnabled(false);
      }
    }
    cons.insets = new Insets(2, SHIFT1, 2, 0);
    panel.add(localeRemoteCheckButtons[0], cons);
    cons.insets = new Insets(2, SHIFT2, 2, 0);
    for (int i = 0; i < 3; i++) {
      cons.gridy++;
      panel.add(localeCheckButtons[i], cons);
    }

    cons.insets = new Insets(8, SHIFT3, 8, 0);
    cons.gridy++;
    panel.add(ngramPanel, cons);

    cons.insets = new Insets(2, SHIFT1, 2, 0);
    cons.gridy++;
    panel.add(localeRemoteCheckButtons[1], cons);
    cons.insets = new Insets(2, SHIFT2, 2, 0);
    for (int i = 0; i < 2; i++) {
      cons.gridy++;
      panel.add(remoteCheckButtons[i], cons);
    }
    cons.insets = new Insets(0, SHIFT3, 0, 0);
    cons.gridy++;
    panel.add(usernameLabel, cons);
    cons.gridy++;
    panel.add(usernameField, cons);
    cons.gridy++;
    panel.add(apiKeyLabel, cons);
    cons.gridy++;
    panel.add(apiKeyField, cons);
    
    cons.insets = new Insets(2, SHIFT2, 2, 0);
    cons.gridy++;
    panel.add(remoteCheckButtons[2], cons);
    cons.insets = new Insets(0, SHIFT3, 0, 0);
    cons.gridy++;
    panel.add(otherServerNameField, cons);
    JLabel serverExampleLabel = new JLabel(" " + WtGeneralTools.getLabel(messages.getString("guiUseServerExample")));
    serverExampleLabel.setEnabled(false);
    cons.gridy++;
    panel.add(serverExampleLabel, cons);
    
    
/*
    JPanel serverPanel = new JPanel();
    serverPanel.setLayout(new GridBagLayout());
    GridBagConstraints cons1 = new GridBagConstraints();
    cons1.insets = new Insets(0, SHIFT2, 0, 0);
    cons1.gridx = 0;
    cons1.gridy = 0;
    cons1.anchor = GridBagConstraints.WEST;
    cons1.fill = GridBagConstraints.NONE;
    cons1.weightx = 0.0f;
    serverPanel.add(useServerBox, cons1);
    cons1.gridx++;
    serverPanel.add(otherServerNameField, cons1);
    JLabel serverExampleLabel = new JLabel(" " + WtGeneralTools.getLabel(messages.getString("guiUseServerExample")));
    serverExampleLabel.setEnabled(false);
    cons1.gridy++;
    serverPanel.add(serverExampleLabel, cons1);
    cons.gridx = 0;
    cons.gridy++;
    panel.add(serverPanel, cons);
*/
/*
    JPanel premiumPanel = new JPanel();
    premiumPanel.setLayout(new GridBagLayout());
    cons1 = new GridBagConstraints();
    cons1.insets = new Insets(0, SHIFT2, 0, 0);
    cons1.gridx = 0;
    cons1.gridy = 0;
    cons1.anchor = GridBagConstraints.WEST;
    cons1.fill = GridBagConstraints.NONE;
    cons1.weightx = 0.0f;
    premiumPanel.add(isPremiumBox, cons1);
    cons1.insets = new Insets(0, SHIFT3, 0, 0);
    cons1.gridy++;
    premiumPanel.add(usernameLabel, cons1);
    cons1.gridy++;
    premiumPanel.add(usernameField, cons1);
    cons1.gridy++;
    premiumPanel.add(apiKeyLabel, cons1);
    cons1.gridy++;
    premiumPanel.add(apiKeyField, cons1);
    cons.gridx = 0;
    cons.gridy++;
    panel.add(premiumPanel, cons);
*/
    paragraphModeBox.setSelected(config.getNumParasToCheck() == 0 || config.onlySingleParagraphMode());
    paragraphModeBox.setEnabled(!config.onlySingleParagraphMode());
    paragraphModeBox.addItemListener(e1 -> {
      if (paragraphModeBox.isSelected()) {
        config.setNumParasToCheck(0);
        config.setUseTextLevelQueue(false);
        config.setMultiThreadLO(false);
        localeCheckButtons[0].setSelected(true);
        for (int i = 1; i < 3; i++) {
          localeCheckButtons[i].setEnabled(false);
        }
      } else {
        config.setNumParasToCheck(-2);
        config.setMultiThreadLO(false);
        if (localeRemoteCheckButtons[0].isSelected()) {
          for (int i = 0; i < 3; i++) {
            localeCheckButtons[i].setEnabled(true);
          }
        }
      }
    });
    saveCacheBox.setSelected(config.saveLoCache() && !config.onlySingleParagraphMode());
    saveCacheBox.setEnabled(!config.onlySingleParagraphMode());
    saveCacheBox.addItemListener(e1 -> {
      config.setSaveLoCache(saveCacheBox.isSelected());
    });
    cons.insets = new Insets(20, SHIFT1, 0, 0);
    cons.gridx = 0;
    cons.gridy++;
    panel.add(paragraphModeBox, cons);
    cons.gridy++;
    panel.add(saveCacheBox, cons);
    return panel;
  }
  
  private void createOfficeElements(GridBagConstraints cons, JPanel portPanel) {

    addOfficeLanguageElements(cons, portPanel);

    cons.gridx = 0;
    cons.gridy++;
    portPanel.add(new JLabel(" "), cons);
    
    cons.gridy++;
    portPanel.add(getMotherTonguePanel(cons), cons);
    
    cons.gridx = 0;
    cons.gridy++;
    portPanel.add(new JLabel(" "), cons);
    
    JCheckBox useLtSpellCheckerBox = new JCheckBox(WtGeneralTools.getLabel(messages.getString("guiUseLtSpellChecker")));
    if (config.onlySingleParagraphMode()) {
      useLtSpellCheckerBox.setEnabled(false);
      useLtSpellCheckerBox.setSelected(false);
    } else {
      useLtSpellCheckerBox.setSelected(config.useLtSpellChecker());
    }
    useLtSpellCheckerBox.addItemListener(e -> {
      config.setUseLtSpellChecker(useLtSpellCheckerBox.isSelected());
    });
    cons.gridy++;
    portPanel.add(useLtSpellCheckerBox, cons);

    JCheckBox useLongMessagesBox = new JCheckBox(WtGeneralTools.getLabel(messages.getString("guiUseLongMessages")));
    useLongMessagesBox.setSelected(config.useLongMessages());
    useLongMessagesBox.addItemListener(e -> {
      config.setUseLongMessages(useLongMessagesBox.isSelected());
    });
    cons.gridy++;
    portPanel.add(useLongMessagesBox, cons);

    JCheckBox markSingleCharBold = new JCheckBox(WtGeneralTools.getLabel(messages.getString("guiMarkSingleCharBold")));
    markSingleCharBold.setSelected(config.markSingleCharBold());
    markSingleCharBold.addItemListener(e -> config.setMarkSingleCharBold(markSingleCharBold.isSelected()));
    cons.gridy++;
    portPanel.add(markSingleCharBold, cons);

    JCheckBox noSynonymsAsSuggestionsBox = new JCheckBox(WtGeneralTools.getLabel(messages.getString("guiNoSynonymsAsSuggestions")));
    noSynonymsAsSuggestionsBox.setSelected(config.noSynonymsAsSuggestions());
    noSynonymsAsSuggestionsBox.addItemListener(e -> {
      config.setNoSynonymsAsSuggestions(noSynonymsAsSuggestionsBox.isSelected());
    });
    cons.gridy++;
    portPanel.add(noSynonymsAsSuggestionsBox, cons);

    JCheckBox includeTrackedChangesBox = new JCheckBox(WtGeneralTools.getLabel(messages.getString("guiIncludeTrackedChanges")));
    includeTrackedChangesBox.setSelected(config.includeTrackedChanges());
    includeTrackedChangesBox.addItemListener(e -> {
      config.setIncludeTrackedChanges(includeTrackedChangesBox.isSelected());
    });
    cons.gridy++;
    portPanel.add(includeTrackedChangesBox, cons);

    JCheckBox enableTmpOffRulesBox = new JCheckBox(WtGeneralTools.getLabel(messages.getString("guiActivateTempOffRules")));
    enableTmpOffRulesBox.setSelected(config.enableTmpOffRules());
    enableTmpOffRulesBox.addItemListener(e -> {
      config.setEnableTmpOffRules(enableTmpOffRulesBox.isSelected());
    });
    cons.gridy++;
    portPanel.add(enableTmpOffRulesBox, cons);

    JCheckBox enableGoalSpecificRulesBox = new JCheckBox(WtGeneralTools.getLabel(messages.getString("guiEnableGoalSpecificRules")));
    enableGoalSpecificRulesBox.setSelected(config.enableGoalSpecificRules());
    enableGoalSpecificRulesBox.addItemListener(e -> {
      config.setEnableGoalSpecificRules(enableGoalSpecificRulesBox.isSelected());
    });
    cons.gridy++;
    portPanel.add(enableGoalSpecificRulesBox, cons);

    JCheckBox filterOverlappingMatchesBox = new JCheckBox(WtGeneralTools.getLabel(messages.getString("guiFilterOverlappingMatches")));
    filterOverlappingMatchesBox.setSelected(config.filterOverlappingMatches());
    filterOverlappingMatchesBox.addItemListener(e -> {
      config.setFilterOverlappingMatches(filterOverlappingMatchesBox.isSelected());
    });
    cons.gridy++;
    portPanel.add(filterOverlappingMatchesBox, cons);

    JCheckBox noCheckGrammarDirectSpeechBox = new JCheckBox(WtGeneralTools.getLabel(messages.getString("guiNoGrammarCheckWithinDirectSpeech")));
    noCheckGrammarDirectSpeechBox.setSelected(config.getCheckDirectSpeech() == WtConfiguration.CHECK_DIRECT_SPEECH_NO);
    noCheckGrammarDirectSpeechBox.addItemListener(e -> {
      config.setCheckDirectSpeech(noCheckGrammarDirectSpeechBox.isSelected() ? 
          WtConfiguration.CHECK_DIRECT_SPEECH_NO : WtConfiguration.CHECK_DIRECT_SPEECH_NO_STYLE);
    });
    
    JCheckBox noCheckStyleDirectSpeechBox = new JCheckBox(WtGeneralTools.getLabel(messages.getString("guiNoStyleCheckWithinDirectSpeech")));
    noCheckStyleDirectSpeechBox.setSelected(config.getCheckDirectSpeech() != WtConfiguration.CHECK_DIRECT_SPEECH_YES);
    noCheckGrammarDirectSpeechBox.setEnabled(noCheckStyleDirectSpeechBox.isSelected());
    noCheckStyleDirectSpeechBox.addItemListener(e -> {
      config.setCheckDirectSpeech(noCheckStyleDirectSpeechBox.isSelected() ? (noCheckGrammarDirectSpeechBox.isSelected() ? 
          WtConfiguration.CHECK_DIRECT_SPEECH_NO : WtConfiguration.CHECK_DIRECT_SPEECH_NO_STYLE) :
          WtConfiguration.CHECK_DIRECT_SPEECH_YES);
      noCheckGrammarDirectSpeechBox.setEnabled(noCheckStyleDirectSpeechBox.isSelected());
    });
    cons.gridy++;
    portPanel.add(noCheckStyleDirectSpeechBox, cons);
    cons.insets = new Insets(0, SHIFT2, 0, 0);

    cons.gridy++;
    portPanel.add(noCheckGrammarDirectSpeechBox, cons);
    cons.insets = new Insets(0, SHIFT1, 0, 0);

    JCheckBox noBackgroundCheckBox = new JCheckBox(WtGeneralTools.getLabel(messages.getString("guiNoBackgroundCheck")));
    noBackgroundCheckBox.setSelected(config.noBackgroundCheck());
    noBackgroundCheckBox.addItemListener(e -> config.setNoBackgroundCheck(noBackgroundCheckBox.isSelected()));
    cons.gridy++;
    portPanel.add(noBackgroundCheckBox, cons);
/*
    cons.gridy++;
    portPanel.add(new JLabel(" "), cons);
    
    cons.gridy++;
    portPanel.add(new JLabel(" "), cons);
   
    addOfficeTextruleElements(cons, portPanel);
*/    
    cons.insets = new Insets(0, SHIFT1, 0, 0);
    cons.gridx = 0;
    cons.gridy++;
    portPanel.add(new JLabel(" "), cons);
    
//    addOfficeTechnicalElements(cons, portPanel);
  }
  
  private int showRemoteServerHint(Component component, boolean otherServer) {
    int ret;
    if(config.useOtherServer() || otherServer) {
        ret = WtOptionPane.showConfirmDialog(component, 
            MessageFormat.format(messages.getString("loRemoteInfoOtherServer"), config.getServerUrl()), 
          messages.getString("loMenuRemoteInfo"), WtOptionPane.OK_CANCEL_OPTION);
    } else {
      ret = WtOptionPane.showConfirmDialog(component, messages.getString("loRemoteInfoDefaultServer"), 
          messages.getString("loMenuRemoteInfo"), WtOptionPane.OK_CANCEL_OPTION);
    }
    return ret;
  }

  @NotNull
  private DefaultTreeModel getTreeModel(DefaultMutableTreeNode rootNode, List<Rule> rules) {
    DefaultTreeModel treeModel = new DefaultTreeModel(rootNode);
    treeModel.addTreeModelListener(new TreeModelListener() {
      @Override
      public void treeNodesChanged(TreeModelEvent e) {
        DefaultMutableTreeNode node = (DefaultMutableTreeNode) e.getTreePath().getLastPathComponent();
        int index = e.getChildIndices()[0];
        node = (DefaultMutableTreeNode) node.getChildAt(index);
        if (node instanceof WtRuleNode) {
          WtRuleNode o = (WtRuleNode) node;
          if (o.getRule().isDefaultOff() || o.getRule().getCategory().isDefaultOff()) {
            if (o.isEnabled()) {
              config.getEnabledRuleIds().add(o.getRule().getId());
              config.getDisabledRuleIds().remove(o.getRule().getId());
            } else {
              config.getEnabledRuleIds().remove(o.getRule().getId());
              config.getDisabledRuleIds().add(o.getRule().getId());
            }
          } else {
            if (o.isEnabled()) {
              config.getDisabledRuleIds().remove(o.getRule().getId());
            } else {
              config.getDisabledRuleIds().add(o.getRule().getId());
            }
          }
          updateProfileRules(rules);
        }
        if (node instanceof WtCategoryNode) {
          WtCategoryNode o = (WtCategoryNode) node;
          if (o.getCategory().isDefaultOff()) {
            if (o.isEnabled()) {
              config.getDisabledCategoryNames().remove(o.getCategory().getName());
              config.getEnabledCategoryNames().add(o.getCategory().getName());
            } else {
              config.getDisabledCategoryNames().add(o.getCategory().getName());
              config.getEnabledCategoryNames().remove(o.getCategory().getName());
            }
          } else {
            if (o.isEnabled()) {
              config.getDisabledCategoryNames().remove(o.getCategory().getName());
            } else {
              config.getDisabledCategoryNames().add(o.getCategory().getName());
            }
          }
        }
      }
      @Override
      public void treeNodesInserted(TreeModelEvent e) {}
      @Override
      public void treeNodesRemoved(TreeModelEvent e) {}
      @Override
      public void treeStructureChanged(TreeModelEvent e) {}
    });
    return treeModel;
  }

  @NotNull
  private MouseAdapter getMouseAdapter() {
    return new MouseAdapter() {
        private void handlePopupEvent(MouseEvent e) {
          JTree tree = (JTree) e.getSource();
          TreePath path = tree.getPathForLocation(e.getX(), e.getY());
          if (path == null) {
            return;
          }
          DefaultMutableTreeNode node
                  = (DefaultMutableTreeNode) path.getLastPathComponent();
          TreePath[] paths = tree.getSelectionPaths();
          boolean isSelected = false;
          if (paths != null) {
            for (TreePath selectionPath : paths) {
              if (selectionPath.equals(path)) {
                isSelected = true;
              }
            }
          }
          if (!isSelected) {
            tree.setSelectionPath(path);
          }
          if (node.isLeaf()) {
            try {
//              int theme = WtDocumentsHandler.getJavaLookAndFeelSet();
//              WtGeneralTools.setJavaLookAndFeel(WtGeneralTools.THEME_SYSTEM);
  
              JPopupMenu popup = new JPopupMenu();
              JMenuItem aboutRuleMenuItem = new JMenuItem(messages.getString("guiAboutRuleMenu"));
              aboutRuleMenuItem.addActionListener(actionEvent -> {
                WtRuleNode node1 = (WtRuleNode) tree.getSelectionPath().getLastPathComponent();
                Rule rule = node1.getRule();
                Language lang = config.getLanguage();
                if(lang == null) {
                  lang = Languages.getLanguageForLocale(Locale.getDefault());
                }
                WtGeneralTools.showRuleInfoDialog(tree, messages.getString("guiAboutRuleTitle"),
                        rule.getDescription(), rule, rule.getUrl(), messages,
                        lang.getShortCodeWithCountryAndVariant());
              });
              popup.add(aboutRuleMenuItem);
              popup.show(tree, e.getX(), e.getY());
//              WtGeneralTools.setJavaLookAndFeel(theme);
            } catch (Exception ex) {
              WtGeneralTools.showErrorMessage(ex);
            }
          }
        }
  
        @Override
        public void mousePressed(MouseEvent e) {
          if (e.isPopupTrigger()) {
            handlePopupEvent(e);
          }
        }
  
        @Override
        public void mouseReleased(MouseEvent e) {
          if (e.isPopupTrigger()) {
            handlePopupEvent(e);
          }
        }
      };
  }

  private boolean getDefaultRuleState(Rule rule) {
    boolean ret = true;
    if ((rule.isDefaultOff() && !rule.isOfficeDefaultOn()) || rule.isOfficeDefaultOff() || rule.getCategory().isDefaultOff()) {
      ret = false;
    }
    return ret;
  }

  @NotNull
  private JPanel getTreeButtonPanel(int num) {
    GridBagConstraints cons;
    JPanel treeButtonPanel = new JPanel();
    cons = new GridBagConstraints();
    cons.gridx = 0;
    cons.gridy = 0;
    cons.weightx = 0;
    cons.weighty = 0;
    JButton selectAllButton = new JButton(messages.getString("guiSelectAll"));
    treeButtonPanel.add(selectAllButton, cons);
    selectAllButton.addActionListener(e -> {
      TreeNode root = (TreeNode) configTree[num].getModel().getRoot();
      for (Enumeration<?> cat = root.children(); cat.hasMoreElements();) {
        WtCategoryNode n = (WtCategoryNode) cat.nextElement();
        n.setEnabled(true);
        for (Enumeration<?> rul = n.children(); rul.hasMoreElements();) {
          WtRuleNode r = (WtRuleNode) rul.nextElement();
          r.setEnabled(true);
        }
      }
      configTree[num].repaint();
    });

    cons.gridx++;
    JButton deselectAllButton = new JButton(messages.getString("guiDeselectAll"));
    treeButtonPanel.add(deselectAllButton, cons);
    deselectAllButton.addActionListener(e -> {
      TreeNode root = (TreeNode) configTree[num].getModel().getRoot();
      for (Enumeration<?> cat = root.children(); cat.hasMoreElements();) {
        WtCategoryNode n = (WtCategoryNode) cat.nextElement();
        n.setEnabled(false);
        for (Enumeration<?> rul = n.children(); rul.hasMoreElements();) {
          WtRuleNode r = (WtRuleNode) rul.nextElement();
          r.setEnabled(false);
        }
      }
      configTree[num].repaint();
    });

    cons.gridx++;
    JButton defaultButton = new JButton(messages.getString("guiDefault"));
    treeButtonPanel.add(defaultButton, cons);
    defaultButton.addActionListener(e -> {
      TreeNode root = (TreeNode) configTree[num].getModel().getRoot();
      for (Enumeration<?> cat = root.children(); cat.hasMoreElements();) {
        WtCategoryNode n = (WtCategoryNode) cat.nextElement();
        n.setEnabled(!n.getCategory().isDefaultOff());
        for (Enumeration<?> rul = n.children(); rul.hasMoreElements();) {
          WtRuleNode r = (WtRuleNode) rul.nextElement();
          r.setEnabled(getDefaultRuleState(r.getRule()));
        }
      }
      configTree[num].repaint();
    });
    
    cons.weightx = 10;
    cons.gridx++;
    treeButtonPanel.add(new JLabel(" "), cons);

    cons.weightx = 0;
    cons.gridx++;
    JButton expandAllButton = new JButton(messages.getString("guiExpandAll"));
    treeButtonPanel.add(expandAllButton, cons);
    expandAllButton.addActionListener(e -> {
      TreeNode root = (TreeNode) configTree[num].getModel().getRoot();
      TreePath parent = new TreePath(root);
      for (Enumeration<?> cat = root.children(); cat.hasMoreElements();) {
        TreeNode n = (TreeNode) cat.nextElement();
        TreePath child = parent.pathByAddingChild(n);
        configTree[num].expandPath(child);
      }
    });
    cons.gridx++;
    JButton collapseAllButton = new JButton(messages.getString("guiCollapseAll"));
    treeButtonPanel.add(collapseAllButton, cons);
    collapseAllButton.addActionListener(e -> {
      TreeNode root = (TreeNode) configTree[num].getModel().getRoot();
      TreePath parent = new TreePath(root);
      for (Enumeration<?> categ = root.children(); categ.hasMoreElements();) {
        TreeNode n = (TreeNode) categ.nextElement();
        TreePath child = parent.pathByAddingChild(n);
        configTree[num].collapsePath(child);
      }
    });
    return treeButtonPanel;
  }
  
  @NotNull
  private JPanel getProfilePanel(List<Rule> rules) {
    profileChanged = true;
    JPanel profilePanel = new JPanel();
    profilePanel.setLayout(new GridBagLayout());
    GridBagConstraints cons = new GridBagConstraints();
    cons.insets = new Insets(4, 4, 0, 8);
    cons.gridx = 0;
    cons.gridy = 0;
    cons.weightx = 1.0f;
    cons.anchor = GridBagConstraints.WEST;
    cons.fill = GridBagConstraints.NONE;
    List<String> profiles = new ArrayList<>();
    String defaultOptions = messages.getString("guiDefaultOptions");
    String userOptions = messages.getString("guiUserProfile");
    profiles.addAll(config.getDefinedProfiles());
    profiles.sort(null);
    profiles.add(0, userOptions);
    String currentProfile = config.getCurrentProfile();
    JComboBox<String> profileBox = new JComboBox<>(profiles.toArray(new String[0]));
    if(currentProfile == null || currentProfile.isEmpty()) {
      profileBox.setSelectedItem(userOptions);
    } else {
      profileBox.setSelectedItem(currentProfile);
    }
    profileBox.addItemListener(e -> {
      if (e.getStateChange() == ItemEvent.SELECTED) {
        if(profileChanged) {
          try {
            //  The configuration has to be saved first to save previous changes
            config.saveConfiguration(null);
            List<String> saveProfiles = new ArrayList<>();
            saveProfiles.addAll(config.getDefinedProfiles());
            if(e.getItem().equals(userOptions)) {
              config.initOptions();
              config.loadConfiguration("");
              config.setCurrentProfile(null);
            } else {
              config.initOptions();
              config.loadConfiguration((String) e.getItem());
              config.setCurrentProfile((String) e.getItem());
            }
            config.addProfiles(saveProfiles);
            restartShow = true;
            dialog.setVisible(false);
          } catch (IOException e1) {
          }
        } else {
          profileChanged = true;
        }
      }
    });
      
    profilePanel.add(new JLabel(addColonToMessageString("guiCurrentProfile")), cons);
    cons.insets = new Insets(6, 16, 0, 8);
    cons.gridy++;
    profilePanel.add(profileBox, cons);
    
    JButton renameButton = new JButton(messages.getString("guiRenameProfile") + "...");
    renameButton.setEnabled(!profileBox.getSelectedItem().equals(defaultOptions) 
        && !profileBox.getSelectedItem().equals(userOptions));
    renameButton.addActionListener(e -> {
      boolean noName = true;
      String profileName = (String) profileBox.getSelectedItem();
      while (noName) {
        profileName = WtOptionPane.showInputDialog(dialog, messages.getString("guiRenameProfile") + ":", profileName);
        if (profileName == null || profileName.equals("")) {
          break;
        }
        profileName = profileName.replaceAll("[ \t=]", "_");
        noName = false;
        while(config.getDefinedProfiles().contains(profileName) || userOptions.equals(profileName)) {
          profileName += "_new";
          noName = true;
        }
      }
      if (profileName != null && !profileName.equals("")) {
        config.removeProfile(config.getCurrentProfile());
        config.addProfile(profileName);
        config.setCurrentProfile(profileName);
        restartShow = true;
        dialog.setVisible(false);
      }
    });
    cons.gridy++;
    profilePanel.add(renameButton, cons);
    
    JButton exportButton = new JButton(messages.getString("guiExportProfile") + "...");
    exportButton.setEnabled(!profileBox.getSelectedItem().equals(defaultOptions) 
        && !profileBox.getSelectedItem().equals(userOptions));
    exportButton.addActionListener(e -> {
      JFileChooser fileChooser = new JFileChooser();
      int choose = fileChooser.showSaveDialog(dialog);
      if (choose == JFileChooser.APPROVE_OPTION) {
        try {
          config.exportProfile((String) profileBox.getSelectedItem(), fileChooser.getSelectedFile());
        } catch (IOException e1) {
        }
      }
    });
    cons.gridx++;
    profilePanel.add(exportButton, cons);
    
    JButton defaultButton = new JButton(defaultOptions);
    defaultButton.addActionListener(e -> {
      List<String> saveProfiles = new ArrayList<>();
      saveProfiles.addAll(config.getDefinedProfiles());
      String saveCurrent = config.getCurrentProfile() == null ? null : config.getCurrentProfile();
      config.initOptions();
      config.addProfiles(saveProfiles);
      config.setCurrentProfile(saveCurrent);
      restartShow = true;
      dialog.setVisible(false);
    });
    cons.gridx = 0;
    cons.gridy++;
    profilePanel.add(defaultButton, cons);
    
    JButton deleteButton = new JButton(messages.getString("guiDeleteProfile"));
    deleteButton.setEnabled(!profileBox.getSelectedItem().equals(defaultOptions) 
        && !profileBox.getSelectedItem().equals(userOptions));
    deleteButton.addActionListener(e -> {
      List<String> saveProfiles = new ArrayList<>();
      saveProfiles.addAll(config.getDefinedProfiles());
      config.initOptions();
      try {
        config.loadConfiguration("");
      } catch (IOException e1) {
      }
      config.setCurrentProfile(null);
      config.addProfiles(saveProfiles);
      config.removeProfile((String)profileBox.getSelectedItem());
      restartShow = true;
      dialog.setVisible(false);
    });
    cons.gridx++;
    profilePanel.add(deleteButton, cons);
    cons.insets = new Insets(16, 4, 0, 8);
    cons.gridx = 0;
    cons.gridy++;
    profilePanel.add(new JLabel(addColonToMessageString("guiAddNewProfile")), cons);
    cons.insets = new Insets(6, 16, 0, 8);
    
    
    JButton addButton = new JButton(messages.getString("guiAddProfile") + "...");
    addButton.addActionListener(e -> {
      boolean noName = true;
      String profileName = "";
      while (noName) {
        profileName = WtOptionPane.showInputDialog(dialog, messages.getString("guiAddNewProfile"), profileName);
        if (profileName == null || profileName.equals("")) {
          break;
        }
        profileName = profileName.replaceAll("[ \t=]", "_");
        noName = false;
        while(config.getDefinedProfiles().contains(profileName) || userOptions.equals(profileName)) {
          profileName += "_new";
          noName = true;
        }
      }
      if (profileName != null && !profileName.equals("")) {
        //  The configuration has to be saved and reloaded first to save previous changes
        try {
          config.saveConfiguration(null);
          config.initOptions();
          config.loadConfiguration(config.getCurrentProfile());
        } catch (IOException e1) {
        }
        config.addProfile(profileName);
        config.setCurrentProfile(profileName);
        profileChanged = false;
        profileBox.addItem(profileName);
        profileBox.setSelectedItem(profileName);
        deleteButton.setEnabled(true);
        renameButton.setEnabled(true);
        exportButton.setEnabled(true);
      }
    });
    cons.gridx = 0;
    cons.gridy++;
    profilePanel.add(addButton, cons);
    
    JButton importButton = new JButton(messages.getString("guiImportProfile") + "...");
    importButton.addActionListener(e -> {
      JFileChooser fileChooser = new JFileChooser();
      int choose = fileChooser.showOpenDialog(dialog);
      if (choose == JFileChooser.APPROVE_OPTION) {
        try {
          //  The configuration has to be saved and reloaded first to save previous changes
          config.saveConfiguration(null);
          config.initOptions();
          config.loadConfiguration(config.getCurrentProfile());
          List<String> saveProfiles = new ArrayList<>();
          saveProfiles.addAll(config.getDefinedProfiles());
          WtConfiguration saveConfig = config.copy(config);
          config.initOptions();
          config.importProfile(fileChooser.getSelectedFile());
          String profileName = config.getCurrentProfile();
          if (profileName != null) {
            config.addProfiles(saveProfiles);
            profileName = profileName.replaceAll("[ \t=]", "_");
            while(config.getDefinedProfiles().contains(profileName) || userOptions.equals(profileName)) {
              profileName += "_new";
            }
            config.setCurrentProfile(profileName);
            config.addProfile(profileName);
            config.saveConfiguration(null);
          } else {
            config.restoreState(saveConfig);;
          }
          restartShow = true;
          dialog.setVisible(false);
        } catch (IOException e1) {
        }
      }
    });
    cons.gridx++;
    profilePanel.add(importButton, cons);
    return profilePanel;
  }
  
  private String addColonToMessageString(String message) {
    String str = messages.getString(message);
    if (!str.endsWith(":")) {
      return str + ":";
    }
    return str;
  }

  @NotNull
  private JPanel getMotherTonguePanel(GridBagConstraints cons) {
    JPanel motherTonguePanel = new JPanel();
    motherTonguePanel.add(new JLabel(messages.getString("guiMotherTongue")), cons);
    JComboBox<String> motherTongueBox = new JComboBox<>(getPossibleLanguages(true));
    if (config.getMotherTongue() != null) {
      motherTongueBox.setSelectedItem(config.getMotherTongue().getTranslatedName(messages));
    }
    motherTongueBox.addItemListener(e -> {
      if (e.getStateChange() == ItemEvent.SELECTED) {
        Language motherTongue;
        if (motherTongueBox.getSelectedItem() instanceof String) {
          motherTongue = getLanguageForLocalizedName(motherTongueBox.getSelectedItem().toString());
        } else {
          motherTongue = (Language) motherTongueBox.getSelectedItem();
        }
        config.setMotherTongue(motherTongue);
      }
    });
    motherTonguePanel.add(motherTongueBox, cons);
    return motherTonguePanel;
  }

  private void addNgramPanel(GridBagConstraints cons, JPanel panel) {
    cons.gridx = 0;
    panel.add(new JLabel((messages.getString("guiNgramDir")) + "  "), cons);
    File dir = config.getNgramDirectory();
    int maxDirDisplayLength = 45;
    String buttonText = dir != null ? StringUtils.abbreviate(dir.getAbsolutePath(), maxDirDisplayLength) : messages.getString("guiNgramDirSelect");
    JButton ngramDirButton = new JButton(buttonText);
    ngramDirButton.addActionListener(e -> {
      File newDir = WtGeneralTools.openDirectoryDialog(owner, dir);
      if (newDir != null) {
        try {
          if (config.getLanguage() != null) {  // may happen in office context
            File checkDir = new File(newDir, config.getLanguage().getShortCode());
            LuceneLanguageModel.validateDirectory(checkDir);
          }
          config.setNgramDirectory(newDir);
          ngramDirButton.setText(StringUtils.abbreviate(newDir.getAbsolutePath(), maxDirDisplayLength));
        } catch (Exception ex) {
          WtGeneralTools.showErrorMessage(ex);
        }
      } else {
        // not the best UI, but this way user can turn off ngram feature without another checkbox
        config.setNgramDirectory(null);
        ngramDirButton.setText(StringUtils.abbreviate(messages.getString("guiNgramDirSelect"), maxDirDisplayLength));
      }
    });
    cons.gridx++;
    panel.add(ngramDirButton, cons);
    JButton helpButton = new JButton(messages.getString("guiNgramHelp"));
    helpButton.addActionListener(e -> WtGeneralTools.openURL("https://dev.languagetool.org/finding-errors-using-n-gram-data"));
    cons.gridx++;
    panel.add(helpButton, cons);
  }

  private String[] getPossibleLanguages(boolean addNoSeletion) {
    List<String> languages = new ArrayList<>();
    if(addNoSeletion) {
      languages.add(NO_SELECTED_LANGUAGE);
    }
    for (Language lang : Languages.get()) {
      languages.add(lang.getTranslatedName(messages));
      languages.sort(null);
    }
    return languages.toArray(new String[0]);
  }

  @Override
  public void actionPerformed(ActionEvent e) {
    if (ACTION_COMMAND_OK.equals(e.getActionCommand())) {
      if (original != null) {
        original.restoreState(config);
      }
      for(JPanel extra : extraPanels) {
        if(extra instanceof WtSavablePanel) {
          ((WtSavablePanel) extra).save();
        }
      }
      if(config.doRemoteCheck() && config.useOtherServer()) {
        String serverName = config.getServerUrl();
        if(serverName == null || (!serverName.startsWith("http://") && !serverName.startsWith("https://"))
            || serverName.endsWith("/") || serverName.endsWith("/v2")) {
          WtOptionPane.showMessageDialog(dialog, WtGeneralTools.getLabel(messages.getString("guiUseServerWarning1")) + "\n" + WtGeneralTools.getLabel(messages.getString("guiUseServerWarning2")));
          if(serverName.endsWith("/")) {
            serverName = serverName.substring(0, serverName.length() - 1);
            config.setOtherServerUrl(serverName);
          }
          if(serverName.endsWith("/v2")) {
            serverName = serverName.substring(0, serverName.length() - 3);
            config.setOtherServerUrl(serverName);
          }
          restartShow = true;
          dialog.setVisible(false);
          return;
        }
      }
      configChanged = true;
      dialog.setVisible(false);
    } else if (ACTION_COMMAND_CANCEL.equals(e.getActionCommand())) {
      dialog.setVisible(false);
    } else if ("Help".equals(e.getActionCommand())) {
      int ind = tabpane.getSelectedIndex();
      int spTab = config.getSpecialTabNames().length;
      if (ind == 0) {
        WtGeneralTools.openURL(WtOfficeTools.getUrl("OptionProfiles"));
      } else if (ind == 1) {
        WtGeneralTools.openURL(WtOfficeTools.getUrl("OptionGeneral"));
      } else if (ind >= 2 && ind <= 3 + spTab) {
        WtGeneralTools.openURL(WtOfficeTools.getUrl("OptionGrammarAndStyle"));
      } else if (ind == 4 + spTab) {
        WtGeneralTools.openURL(WtOfficeTools.getUrl("OptionDefaultColors"));
      } else if (ind == 5 + spTab) {
        WtGeneralTools.openURL(WtOfficeTools.getUrl("OptionTechnicalSettings"));
      } else if (ind == 6 + spTab) {
        WtGeneralTools.openURL(WtOfficeTools.getUrl("OptionAiSupport"));
      }
    }
  }

  /**
   * Get the Language object for the given localized language name.
   * 
   * @param languageName e.g. <code>English</code> or <code>German</code> (case is significant)
   * @return a Language object or <code>null</code> if the language could not be found
   */
  @Nullable
  private Language getLanguageForLocalizedName(String languageName) {
    for (Language element : Languages.get()) {
      if (languageName.equals(element.getTranslatedName(messages))) {
        return element;
      }
    }
    return null;
  }

  static class CategoryComparator implements Comparator<Rule> {

    @Override
    public int compare(Rule r1, Rule r2) {
      boolean hasCat = r1.getCategory() != null && r2.getCategory() != null;
      if (hasCat) {
        int res = r1.getCategory().getName().compareTo(r2.getCategory().getName());
        if (res == 0) {
          return r1.getDescription() != null && r2.getDescription() != null ? r1.getDescription().compareToIgnoreCase(r2.getDescription()) : 0;
        }
        return res;
      }
      return r1.getDescription() != null && r2.getDescription() != null ? r1.getDescription().compareToIgnoreCase(r2.getDescription()) : 0;
    }

  }
  
  /**
   * Update display of rules tree
   */
  private void updateRulesTrees(List<Rule> rules) {
    String[] specialTabNames = config.getSpecialTabNames();
    int numConfigTrees = 2 + specialTabNames.length;
    for (int i = 0; i < numConfigTrees; i++) {
      if(i == 0) {
        rootNode[i] = createTree(rules, false, null, rootNode[i]);   //  grammar options
      } else if(i == 1) {
        rootNode[i] = createTree(rules, true, null, rootNode[i]);    //  Style options
      } else {
        rootNode[i] = createTree(rules, true, specialTabNames[i - 2], rootNode[i]);    //  Special tab options
      }
      configTree[i].setModel(getTreeModel(rootNode[i], rules));
    }
  }
  
  /**
   * Update display of profile rules
   */
  private void updateProfileRules(List<Rule> rules) {
    getChangedRulesPanel(rules, false, disabledRulesPanel);
    getChangedRulesPanel(rules, true , enabledRulesPanel);
  }

  
  /** Panel to select disabled default rules
   * @since 5.4
   */
  private JPanel getChangedRulesPanel(List<Rule> rules, boolean enabledRules, JPanel panel) {
    if (panel == null) {
      panel = new JPanel();
    } else {
      panel.removeAll();
    }
//    panel.setBackground(new Color(169,169,169));
    panel.setBorder(BorderFactory.createLineBorder(Color.black));
    panel.setLayout(new GridBagLayout());
    GridBagConstraints cons = new GridBagConstraints();
    cons.gridx = 0;
    cons.gridy = 0;
    cons.weightx = 1.0f;
    cons.anchor = GridBagConstraints.WEST;
    cons.fill = GridBagConstraints.NONE;
    cons.insets = new Insets(4, 3, 0, 4);
    
    List<String> changedRuleIds;
    if (enabledRules) {
      changedRuleIds = new ArrayList<String>(config.getEnabledRuleIds());
    } else {
      changedRuleIds = new ArrayList<String>(config.getDisabledRuleIds());
    }
    
    if (changedRuleIds != null) {
      List<JCheckBox> ruleCheckboxes = new ArrayList<>();
      for (int i = changedRuleIds.size() - 1; i >= 0; i--) {
        String ruleId = changedRuleIds.get(i);
        String ruleDescription = null;
        for (Rule rule : rules) {
          if (rule.getId().equals(ruleId)) {
            if ((enabledRules && (rule.getCategory().isDefaultOff() || (rule.isDefaultOff() && !rule.isOfficeDefaultOn()))) ||
                (!enabledRules && !rule.getCategory().isDefaultOff() && (!rule.isDefaultOff() || rule.isOfficeDefaultOn()))) {
              ruleDescription = rule.getDescription();
            } else {
              if (enabledRules) {
                config.removeEnabledRuleId(ruleId);
              } else {
                config.removeDisabledRuleId(ruleId);
              }
            }
            
            break;
          }
        }
        if (ruleDescription != null) {
          JCheckBox ruleCheckbox = new JCheckBox(ruleDescription);
          ruleCheckbox.setName(ruleId);
          ruleCheckboxes.add(ruleCheckbox);
          ruleCheckbox.setSelected(enabledRules);
          panel.add(ruleCheckbox, cons);
          ruleCheckbox.addActionListener(e -> {
            if (ruleCheckbox.isSelected()) {
              config.getEnabledRuleIds().add(ruleCheckbox.getName());
              config.getDisabledRuleIds().remove(ruleCheckbox.getName());
              updateRulesTrees(rules);
            } else {
              config.getEnabledRuleIds().remove(ruleCheckbox.getName());
              config.getDisabledRuleIds().add(ruleCheckbox.getName());
              updateRulesTrees(rules);
            }
          });
          cons.gridx = 0;
          cons.gridy++;
        }
      }
    }
    return panel;
  }
  
  private String[] getUnderlineTypes() {
    String[] types = {
      messages.getString("guiUTypeWave"),
      messages.getString("guiUTypeBoldWave"),
      messages.getString("guiUTypeBold"),
      messages.getString("guiUTypeDash")};
    return types;
  }

  private int getUnderlineType(String category, String ruleId) {
    short nType = config.getUnderlineType(category, ruleId);
    if (nType == WtConfiguration.UNDERLINE_BOLDWAVE) {
      return 1;
    } else if (nType == WtConfiguration.UNDERLINE_BOLD) {
      return 2;
    } else if (nType == WtConfiguration.UNDERLINE_DASH) {
      return 3;
    } else {
      return 0;
    }
  }

  private void setUnderlineType(int index, String category, String ruleId) {
    if (ruleId == null) {
      if (index == 1) {
        config.setUnderlineType(category, WtConfiguration.UNDERLINE_BOLDWAVE);
      } else if (index == 2) {
        config.setUnderlineType(category, WtConfiguration.UNDERLINE_BOLD);
      } else if (index == 3) {
        config.setUnderlineType(category, WtConfiguration.UNDERLINE_DASH);
      } else {
        config.setDefaultUnderlineType(category);
      }
    } else {
      if (index == 1) {
        config.setUnderlineRuleType(ruleId, WtConfiguration.UNDERLINE_BOLDWAVE);
      } else if (index == 2) {
        config.setUnderlineRuleType(ruleId, WtConfiguration.UNDERLINE_BOLD);
      } else if (index == 3) {
        config.setUnderlineRuleType(ruleId, WtConfiguration.UNDERLINE_DASH);
      } else {
        config.setDefaultUnderlineRuleType(ruleId);
      }
    }
  }

  /**  Panel to choose underline Colors
   *   @since 4.2
   *//*
  private JPanel getUnderlineColorPanel(List<Rule> rules) {
    JPanel panel = new JPanel();

    panel.setLayout(new GridBagLayout());
    GridBagConstraints cons = new GridBagConstraints();
    cons.gridx = 0;
    cons.gridy = 0;
    cons.weightx = 0.0f;
    cons.fill = GridBagConstraints.NONE;
    cons.anchor = GridBagConstraints.NORTHWEST;

    List<String> categories = new ArrayList<>();
    List<Boolean> isDefault = new ArrayList<>();
    for (Rule rule : rules) {
      String category = rule.getCategory().getName();
      boolean contain = false;
      for(String c : categories) {
        if (c.equals(category)) {
          contain = true;
          break;
        }
      }
      if (!contain) {
        categories.add(category);
        isDefault.add(!rule.getCategory().isDefaultOff());
      }
    }
    List<JLabel> categoryLabel = new ArrayList<>();
    List<JLabel> underlineLabel = new ArrayList<>();
    List<JButton> changeButton = new ArrayList<>();
    List<JButton> defaultButton = new ArrayList<>();
    List<JComboBox<String>> underlineType  = new ArrayList<>();
    for(int nCat = 0; nCat < categories.size(); nCat++) {
      categoryLabel.add(new JLabel(categories.get(nCat) + " "));
      underlineLabel.add(new JLabel(" \u2588\u2588\u2588 "));  // \u2587 is smaller
      underlineLabel.get(nCat).setForeground(config.getUnderlineColor(categories.get(nCat), null, isDefault.get(nCat)));
      underlineLabel.get(nCat).setBackground(config.getUnderlineColor(categories.get(nCat), null, isDefault.get(nCat)));
      JLabel uLabel = underlineLabel.get(nCat);
      String cLabel = categories.get(nCat);
      boolean iLabel = isDefault.get(nCat);
      panel.add(categoryLabel.get(nCat), cons);

      underlineType.add(new JComboBox<>(getUnderlineTypes()));
      JComboBox<String> uLineType = underlineType.get(nCat);
      uLineType.setSelectedIndex(getUnderlineType(cLabel, null));
      uLineType.addItemListener(e -> {
        if (e.getStateChange() == ItemEvent.SELECTED) {
          setUnderlineType(uLineType.getSelectedIndex(), cLabel, null);
        }
      });
      cons.gridx++;
      panel.add(uLineType, cons);
      cons.gridx++;
      panel.add(underlineLabel.get(nCat), cons);

      changeButton.add(new JButton(messages.getString("guiUColorChange")));
      changeButton.get(nCat).addActionListener(e -> {
        Color oldColor = uLabel.getForeground();
        dialog.setAlwaysOnTop(false);
        
        WtColorChooser colorChooser = new WtColorChooser(oldColor);
        ActionListener okActionListener = new ActionListener() {
          public void actionPerformed(ActionEvent actionEvent) {
            Color newColor = colorChooser.getColor();
            if(newColor != null && newColor != oldColor) {
              uLabel.setForeground(newColor);
              config.setUnderlineColor(cLabel, newColor);
            }
            dialog.setAlwaysOnTop(true);
          }
        };
        // For cancel selection, change button background to red
        ActionListener cancelActionListener = new ActionListener() {
          public void actionPerformed(ActionEvent actionEvent) {
            dialog.setAlwaysOnTop(true);
          }
        };
        JDialog colorDialog = WtColorChooser.createDialog(dialog, messages.getString("guiUColorDialogHeader"), true,
            colorChooser, okActionListener, cancelActionListener);
        colorDialog.setAlwaysOnTop(true);
        colorDialog.toFront();
        colorDialog.setVisible(true);
      });
      cons.gridx++;
      panel.add(changeButton.get(nCat), cons);
  
      defaultButton.add(new JButton(messages.getString("guiUColorDefault")));
      defaultButton.get(nCat).addActionListener(e -> {
        config.setDefaultUnderlineColor(cLabel);
        uLabel.setForeground(config.getUnderlineColor(cLabel, null, iLabel));
        config.setDefaultUnderlineType(cLabel);
        uLineType.setSelectedIndex(getUnderlineType(cLabel, null));
      });
      cons.gridx++;
      panel.add(defaultButton.get(nCat), cons);
      cons.gridx = 0;
      cons.gridy++;
    }
    
    return panel;
  }
*/
  /**  Panel to choose underline Colors
   *   and rule options (if exists)
   *   @since 5.3
   */
  @NotNull
  private JPanel getRuleOptionsPanel(int num) {
    category = "";
    rule = null;
    JPanel ruleOptionsPanel = new JPanel();
    ruleOptionsPanel.setLayout(new GridBagLayout());
    GridBagConstraints cons0 = new GridBagConstraints();
    cons0.gridx = 0;
    cons0.gridy = 0;
    cons0.fill = GridBagConstraints.NONE;
    cons0.anchor = GridBagConstraints.NORTHWEST;
    cons0.weightx = 2.0f;
    cons0.weighty = 0.0f;
    cons0.insets = new Insets(3, 8, 3, 0);
    ruleOptionsPanel.setBorder(BorderFactory.createLineBorder(Color.black));
    
    //  Color Panel
    JPanel colorPanel = new JPanel();
    colorPanel.setLayout(null);
    colorPanel.setBounds(0, 0, 120, 10);

    colorPanel.setLayout(new GridBagLayout());
    GridBagConstraints cons1 = new GridBagConstraints();
    cons1.insets = new Insets(0, 0, 0, 0);
    cons1.gridx = 0;
    cons1.gridy = 0;
    cons1.weightx = 0.0f;
    cons1.fill = GridBagConstraints.NONE;
    cons1.anchor = GridBagConstraints.NORTHWEST;

    JLabel underlineStyle = new JLabel(messages.getString("guiUColorStyleLabel") + " ");
    colorPanel.add(underlineStyle);

    JLabel underlineLabel = new JLabel(COLOR_LABEL);

    JComboBox<String> underlineType = new JComboBox<>(getUnderlineTypes());
    underlineType.setEnabled(!config.onlySingleParagraphMode());
    underlineType.setSelectedIndex(getUnderlineType(category, (rule == null ? null : rule.getId())));
    underlineType.addItemListener(e -> {
      if (e.getStateChange() == ItemEvent.SELECTED) {
        setUnderlineType(underlineType.getSelectedIndex(), category, (rule == null ? null : rule.getId()));
      }
    });
    cons1.gridx++;
    colorPanel.add(underlineType);
    cons1.gridx++;
    colorPanel.add(underlineLabel);

    JButton changeButton = new JButton(messages.getString("guiUColorChange"));
    changeButton.setEnabled(!config.onlySingleParagraphMode());
    changeButton.addActionListener(e -> {
      Color oldColor = underlineLabel.getForeground();
      dialog.setAlwaysOnTop(false);
      try {
//        int theme = WtDocumentsHandler.getJavaLookAndFeelSet();
//        WtGeneralTools.setJavaLookAndFeel(WtGeneralTools.THEME_SYSTEM);
        JColorChooser colorChooser = new JColorChooser(oldColor);
        ActionListener okActionListener = new ActionListener() {
          public void actionPerformed(ActionEvent actionEvent) {
            Color newColor = colorChooser.getColor();
            if(newColor != null && newColor != oldColor) {
              underlineLabel.setForeground(newColor);
              if (rule == null) {
                config.setUnderlineColor(category, newColor);
              } else {
                config.setUnderlineRuleColor(rule.getId(), newColor);
              }
            }
            dialog.setAlwaysOnTop(true);
          }
        };
        // For cancel selection, change button background to red
        ActionListener cancelActionListener = new ActionListener() {
          public void actionPerformed(ActionEvent actionEvent) {
            dialog.setAlwaysOnTop(true);
          }
        };
        JDialog colorDialog = JColorChooser.createDialog(dialog, messages.getString("guiUColorDialogHeader"), true,
            colorChooser, okActionListener, cancelActionListener);
        colorDialog.setAlwaysOnTop(true);
        colorDialog.toFront();
        colorDialog.setVisible(true);
//        WtGeneralTools.setJavaLookAndFeel(theme);
      } catch (Exception e1) {
        WtMessageHandler.printException(e1);
      }
    });
    cons1.gridx++;
    colorPanel.add(changeButton);
  
    JButton defaultButton = new JButton(messages.getString("guiUColorDefault"));
    defaultButton.setEnabled(!config.onlySingleParagraphMode());
    defaultButton.addActionListener(e -> {
      String ruleId = (rule == null ? null : rule.getId());
      if (rule == null) {
        config.setDefaultUnderlineColor(category);
      } else {
        config.setDefaultUnderlineRuleColor(ruleId);
      }
      underlineLabel.setForeground(config.getUnderlineColor(category, ruleId));
      if ( rule == null) {
        config.setDefaultUnderlineType(category);
      } else {
        config.setDefaultUnderlineRuleType(ruleId);
      }
      underlineType.setSelectedIndex(getUnderlineType(category, ruleId));
      config.removeConfigurableValue(ruleId);
    });
    cons1.gridx++;
    colorPanel.add(defaultButton);
    colorPanel.setVisible(false);
    // End of Color Panel
    
    List<JPanel> specialOptionPanels = new ArrayList<>();
    
    ruleOptionsPanel.add(colorPanel, cons0);
    cons0.gridx = 0;
    cons0.gridy = 1;
    
    configTree[num].addTreeSelectionListener(e -> {
      DefaultMutableTreeNode node = (DefaultMutableTreeNode)
          configTree[num].getLastSelectedPathComponent();
      if (node != null) {
        if (specialOptionPanels.size() > 0) {
          for (JPanel optionPanel : specialOptionPanels) {
            optionPanel.setVisible(false);
            ruleOptionsPanel.remove(optionPanel);
          }
          specialOptionPanels.clear();
        }
        ruleOptionsPanel.setVisible(false);
        if (node instanceof WtRuleNode) {
          WtRuleNode o = (WtRuleNode) node;
          rule = o.getRule();
          category = rule.getCategory().getName();
          String ruleId = rule.getId();
          underlineLabel.setForeground(config.getUnderlineColor(category, ruleId));
          underlineLabel.setBackground(config.getUnderlineColor(category, ruleId));
          underlineType.setSelectedIndex(getUnderlineType(category, ruleId));
          colorPanel.setVisible(true);
          RuleOption[] ruleOptions = rule.getRuleOptions();
          if (ruleOptions != null && ruleOptions.length > 0) {
            Object[] obj = new Object[ruleOptions.length];
            for (int i = 0; i < ruleOptions.length; i++) {
              // Start of special option panel
              JPanel specialOptionPanel = new JPanel();
              specialOptionPanels.add(specialOptionPanel);
              specialOptionPanel.setLayout(new GridBagLayout());
              GridBagConstraints cons2 = new GridBagConstraints();
              cons2.gridx = 0;
              cons2.gridy = 0;
              cons2.weightx = 2.0f;
              cons2.anchor = GridBagConstraints.WEST;
              RuleOption ruleOption = ruleOptions[i];
              int n = i;

              Object defValue = ruleOption.getDefaultValue();
              
              if (defValue instanceof Boolean) {
                JCheckBox isTrueBox = new JCheckBox(ruleOption.getConfigureText());
                boolean value = config.getConfigValueByID(rule.getId(), i, Boolean.class, (Boolean) defValue);
                isTrueBox.setSelected(value);
                obj[n] = value;
                isTrueBox.addItemListener(e1 -> {
                  obj[n] = isTrueBox.isSelected();
                  config.setConfigurableValue(rule.getId(), obj);
                });
                specialOptionPanel.add(isTrueBox, cons2);
              } else {
                JLabel ruleLabel = new JLabel(ruleOption.getConfigureText() + " ");
                specialOptionPanel.add(ruleLabel, cons2);
    
                cons2.gridx++;
                JTextField ruleValueField = new JTextField("   ", 3);
                ruleValueField.setMinimumSize(new Dimension(50, 28));  // without this the box is just a few pixels small, but why?
                String fieldValue;
                if (defValue instanceof Integer) {
                  obj[n] = (int) config.getConfigValueByID(rule.getId(), i, Integer.class, (Integer) defValue);
                  fieldValue = Integer.toString((int) obj[n]);
                } else if (defValue instanceof Character) {
                  obj[n] = (char) config.getConfigValueByID(rule.getId(), i, Character.class, (Character) defValue);
                  fieldValue = Character.toString((char) obj[n]);
                } else if (defValue instanceof Double) {
                  obj[n] = (double) config.getConfigValueByID(rule.getId(), i, Double.class, (Double) defValue);
                  fieldValue = Double.toString((double) obj[n]);
                } else if (defValue instanceof Float) {
                  obj[n] = (float) config.getConfigValueByID(rule.getId(), i, Float.class, (Float) defValue);
                  fieldValue = Float.toString((float) obj[n]);
                } else {
                  obj[n] = (String) config.getConfigValueByID(rule.getId(), i, String.class, (String) defValue);
                  fieldValue = (String) obj[n];
                }
                ruleValueField.setText(fieldValue);
                specialOptionPanel.add(ruleValueField, cons2);
    
                ruleValueField.getDocument().addDocumentListener(new DocumentListener() {
                  @Override
                  public void insertUpdate(DocumentEvent e) {
                    changedUpdate(e);
                  }
    
                  @Override
                  public void removeUpdate(DocumentEvent e) {
                    changedUpdate(e);
                  }
    
                  @Override
                  public void changedUpdate(DocumentEvent e) {
                    try {
                      if (rule != null) {
                        RuleOption[] ruleOptions = rule.getRuleOptions();
                        if (ruleOptions != null && ruleOptions.length > 0) {
                          boolean isCorrect = false;
                          if (defValue instanceof Integer) {
                            int num = Integer.parseInt(ruleValueField.getText());
                            if (num < (int) ruleOption.getMinConfigurableValue()) {
                              num = (int) ruleOption.getMinConfigurableValue();
                              ruleValueField.setForeground(Color.RED);
                            } else if (num > (int) ruleOption.getMaxConfigurableValue()) {
                              num = (int) ruleOption.getMaxConfigurableValue();
                              ruleValueField.setForeground(Color.RED);
                            } else {
                              ruleValueField.setForeground(null);
                              isCorrect = true;
                              obj[n] = num;
                            }
                          } else if (defValue instanceof Character) {
                            char num = ruleValueField.getText().charAt(0);
                            if (num < (char) ruleOption.getMinConfigurableValue()) {
                              num = (char) ruleOption.getMinConfigurableValue();
                              ruleValueField.setForeground(Color.RED);
                            } else if (num > (char) ruleOption.getMaxConfigurableValue()) {
                              num = (char) ruleOption.getMaxConfigurableValue();
                              ruleValueField.setForeground(Color.RED);
                            } else {
                              ruleValueField.setForeground(null);
                              isCorrect = true;
                              obj[n] = num;
                            }
                          } else if (defValue instanceof Double) {
                            double num = Double.parseDouble(ruleValueField.getText());
                            if (num < (double) ruleOption.getMinConfigurableValue()) {
                              num = (double) ruleOption.getMinConfigurableValue();
                              ruleValueField.setForeground(Color.RED);
                            } else if (num > (double) ruleOption.getMaxConfigurableValue()) {
                              num = (double) ruleOption.getMaxConfigurableValue();
                              ruleValueField.setForeground(Color.RED);
                            } else {
                              ruleValueField.setForeground(null);
                              isCorrect = true;
                              obj[n] = num;
                            }
                          } else if (defValue instanceof Float) {
                            float num = Float.parseFloat(ruleValueField.getText());
                            if (num < (float) ruleOption.getMinConfigurableValue()) {
                              num = (float) ruleOption.getMinConfigurableValue();
                              ruleValueField.setForeground(Color.RED);
                            } else if (num > (float) ruleOption.getMaxConfigurableValue()) {
                              num = (float) ruleOption.getMaxConfigurableValue();
                              ruleValueField.setForeground(Color.RED);
                            } else {
                              ruleValueField.setForeground(null);
                              isCorrect = true;
                              obj[n] = num;
                            }
                          } else {
                            String num = ruleValueField.getText();
                            ruleValueField.setForeground(null);
                            isCorrect = true;
                            obj[n] = num;
                          }
                          if (isCorrect) {
                            config.setConfigurableValue(rule.getId(), obj);
                          }
                        }
                      }
                    } catch (Exception ex) {
                      ruleValueField.setForeground(Color.RED);
                    }
                  }
                });
              }
              ruleOptionsPanel.add(specialOptionPanel, cons0);
              ruleOptionsPanel.setBorder(BorderFactory.createLineBorder(Color.black));
              ruleOptionsPanel.setVisible(true);
              cons0.gridy++;
              // End of special option panel
            }
          }
        } else if (node instanceof WtCategoryNode) {
          WtCategoryNode o = (WtCategoryNode) node;
          category = o.getCategory().getName();
          underlineLabel.setForeground(config.getUnderlineColor(category, null));
          underlineLabel.setBackground(config.getUnderlineColor(category, null));
          underlineType.setSelectedIndex(getUnderlineType(category, null));
          colorPanel.setVisible(true);
          rule = null;
        }
        ruleOptionsPanel.setVisible(true);
      }
    });
    return ruleOptionsPanel;
  }
  
  private JPanel getColorPanel(String category, String ruleId) {
    //  Color Panel
    JPanel colorPanel = new JPanel();
    colorPanel.setLayout(null);
    colorPanel.setBounds(0, 0, 120, 10);

    colorPanel.setLayout(new GridBagLayout());
    GridBagConstraints cons1 = new GridBagConstraints();
    cons1.insets = new Insets(0, 0, 0, 0);
    cons1.gridx = 0;
    cons1.gridy = 0;
    cons1.weightx = 0.0f;
    cons1.fill = GridBagConstraints.NONE;
    cons1.anchor = GridBagConstraints.NORTHWEST;

    JLabel underlineStyle = new JLabel(messages.getString("guiUColorStyleLabel") + " ");
    colorPanel.add(underlineStyle);

    JLabel underlineLabel = new JLabel(COLOR_LABEL);

    JComboBox<String> underlineType = new JComboBox<>(getUnderlineTypes());
    underlineType.setSelectedIndex(getUnderlineType(category, ruleId));
    underlineType.addItemListener(e -> {
      if (e.getStateChange() == ItemEvent.SELECTED) {
        setUnderlineType(underlineType.getSelectedIndex(), category, ruleId);
      }
    });
    cons1.gridx++;
    colorPanel.add(underlineType);
    cons1.gridx++;
    colorPanel.add(underlineLabel);

    JButton changeButton = new JButton(messages.getString("guiUColorChange"));
    changeButton.addActionListener(e -> {
      Color oldColor = underlineLabel.getForeground();
      dialog.setAlwaysOnTop(false);
      try {
//        int theme = WtDocumentsHandler.getJavaLookAndFeelSet();
//        WtGeneralTools.setJavaLookAndFeel(WtGeneralTools.THEME_SYSTEM);
        JColorChooser colorChooser = new JColorChooser(oldColor);
        ActionListener okActionListener = new ActionListener() {
          public void actionPerformed(ActionEvent actionEvent) {
            Color newColor = colorChooser.getColor();
            if(newColor != null && newColor != oldColor) {
              underlineLabel.setForeground(newColor);
              if (ruleId == null) {
                config.setUnderlineColor(category, newColor);
              } else {
                config.setUnderlineRuleColor(ruleId, newColor);
              }
            }
            dialog.setAlwaysOnTop(true);
          }
        };
        // For cancel selection, change button background to red
        ActionListener cancelActionListener = new ActionListener() {
          public void actionPerformed(ActionEvent actionEvent) {
            dialog.setAlwaysOnTop(true);
          }
        };
        JDialog colorDialog = JColorChooser.createDialog(dialog, messages.getString("guiUColorDialogHeader"), true,
            colorChooser, okActionListener, cancelActionListener);
        colorDialog.setAlwaysOnTop(true);
        colorDialog.toFront();
        colorDialog.setVisible(true);
//        WtGeneralTools.setJavaLookAndFeel(theme);
      } catch (Exception e1) {
        WtMessageHandler.printException(e1);
      }
    });
    cons1.gridx++;
    colorPanel.add(changeButton);
  
    JButton defaultButton = new JButton(messages.getString("guiUColorDefault"));
    defaultButton.addActionListener(e -> {
      if (ruleId == null) {
        config.setDefaultUnderlineColor(category);
      } else {
        config.setDefaultUnderlineRuleColor(ruleId);
      }
      underlineLabel.setForeground(config.getUnderlineColor(category, ruleId));
      if ( ruleId == null) {
        config.setDefaultUnderlineType(category);
      } else {
        config.setDefaultUnderlineRuleType(ruleId);
      }
      underlineType.setSelectedIndex(getUnderlineType(category, ruleId));
      config.removeConfigurableValue(ruleId);
    });
    cons1.gridx++;
    colorPanel.add(defaultButton);
    underlineLabel.setForeground(config.getUnderlineColor(category, ruleId));
    underlineLabel.setBackground(config.getUnderlineColor(category, ruleId));
    underlineType.setSelectedIndex(getUnderlineType(category, ruleId));
    colorPanel.setVisible(true);
    // End of Color Panel
    return colorPanel;
  }
  
  private JPanel getAiInstallationHint() {
    JPanel panel = new JPanel();
    panel.setLayout(new GridBagLayout());
    GridBagConstraints cons = new GridBagConstraints();
    cons.insets = new Insets(2, 2, 2, 2);
    cons.gridx = 0;
    cons.gridy = 0;
    cons.weightx = 1.0f;
    cons.fill = GridBagConstraints.HORIZONTAL;
    cons.anchor = GridBagConstraints.NORTHWEST;
    JLabel label = new JLabel(messages.getString("guiAiInstallationHint") + ": ");
    panel.add(label, cons);
    cons.gridx++;
    cons.weightx = 1.0f;
    JButton installButton = new JButton(messages.getString("guiAiInstallationButton"));
    installButton.addActionListener(e -> {
      WtGeneralTools.openURL(WtOfficeTools.getUrl("localAiInstallation"));
    });
    panel.add(installButton, cons);
    return panel;
  }
  
  private JPanel getOfficeAiTextElements() {
    JPanel aiOptionPanel = new JPanel();
    aiOptionPanel.setLayout(new GridBagLayout());
    GridBagConstraints cons = new GridBagConstraints();
    cons.insets = new Insets(6, 6, 6, 6);
    cons.gridx = 0;
    cons.gridy = 0;
    cons.anchor = GridBagConstraints.NORTHWEST;
    cons.fill = GridBagConstraints.HORIZONTAL;
    cons.weightx = 10.0f;
    cons.weighty = 0.0f;
    
    JLabel otherUrlLabel = new JLabel(messages.getString("guiAiUrl") + ":");

    JTextField aiUrlField = new JTextField(config.aiUrl() ==  null ? "" : config.aiUrl(), 25);
    aiUrlField.setMinimumSize(new Dimension(100, 25));
    aiUrlField.getDocument().addDocumentListener(new DocumentListener() {
      @Override
      public void insertUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void removeUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void changedUpdate(DocumentEvent e) {
        String serverName = aiUrlField.getText();
        serverName = serverName.trim();
        if(serverName.isEmpty()) {
          serverName = null;
        }
        if (config.isValidAiServerUrl(serverName)) {
          aiUrlField.setForeground(null);
          config.setAiUrl(serverName);;
        } else {
          aiUrlField.setForeground(Color.RED);
        }
      }
    });

    JLabel modelLabel = new JLabel(messages.getString("guiAiModel") + ":");

    JTextField modelField = new JTextField(config.aiModel() ==  null ? "" : config.aiModel(), 25);
    modelField.setMinimumSize(new Dimension(100, 25));
    modelField.getDocument().addDocumentListener(new DocumentListener() {
      @Override
      public void insertUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void removeUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void changedUpdate(DocumentEvent e) {
        String model = modelField.getText();
        model = model.trim();
        if(model.isEmpty()) {
          model = null;
        }
        if (model != null) {
          config.setAiModel(model);
        }
      }
    });

    JLabel apiKeyLabel = new JLabel(messages.getString("guiAiApiKey") + ":");

    JTextField apiKeyField = new JTextField(config.aiApiKey() ==  null ? "" : config.aiApiKey(), 25);
    apiKeyField.setMinimumSize(new Dimension(100, 25));
    apiKeyField.getDocument().addDocumentListener(new DocumentListener() {
      @Override
      public void insertUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void removeUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void changedUpdate(DocumentEvent e) {
        String apiKey = apiKeyField.getText();
        apiKey = apiKey.trim();
        if(apiKey.isEmpty()) {
          apiKey = null;
        }
        if (apiKey != null) {
          config.setAiApiKey(apiKey);
        }
      }
    });
    
    JRadioButton[] radioButtons = new JRadioButton[3];
    ButtonGroup showStylisticChangesGroup = new ButtonGroup();
    radioButtons[0] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiAiShowNoStylisticChanges")));
    radioButtons[0].addActionListener(e -> {
      config.setAiShowStylisticChanges(0);
    });
    radioButtons[1] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiAiShowSmallStylisticChanges")));
    radioButtons[1].addActionListener(e -> {
      config.setAiShowStylisticChanges(1);
    });
    radioButtons[2] = new JRadioButton(WtGeneralTools.getLabel(messages.getString("guiAiShowAllStylisticChanges")));
    radioButtons[2].addActionListener(e -> {
      config.setAiShowStylisticChanges(2);
    });

    for (int i = 0; i < 3; i++) {
      if (config.aiShowStylisticChanges() == i) {
        radioButtons[i].setSelected(true);
      } else {
        radioButtons[i].setSelected(false);
      }
      showStylisticChangesGroup.add(radioButtons[i]);
    }
    
    JCheckBox aiAutoSuggestionBox = new JCheckBox(messages.getString("guiAiShowStylisticSuggestion"));
    aiAutoSuggestionBox.setSelected(config.aiAutoSuggestion());
    aiAutoSuggestionBox.addItemListener(e -> {
      config.setAiAutoSuggestion(aiAutoSuggestionBox.isSelected());
    });

    JCheckBox autoCorrectBox = new JCheckBox(messages.getString("guiAiAutoCorrect"));
    autoCorrectBox.setSelected(config.aiAutoCorrect());
    autoCorrectBox.addItemListener(e -> {
      config.setAiAutoCorrect(autoCorrectBox.isSelected());
      for (JRadioButton rButton : radioButtons) {
        rButton.setEnabled(autoCorrectBox.isSelected());
      }
      aiAutoSuggestionBox.setEnabled(autoCorrectBox.isSelected());
    });

    JCheckBox useAiSupportBox = new JCheckBox(messages.getString("guiUseAiSupport"));
    useAiSupportBox.setSelected(config.useAiSupport());
    useAiSupportBox.addItemListener(e -> {
      config.setUseAiSupport(useAiSupportBox.isSelected());
      aiUrlField.setEnabled(useAiSupportBox.isSelected());
      modelField.setEnabled(useAiSupportBox.isSelected());
      apiKeyField.setEnabled(useAiSupportBox.isSelected());
      autoCorrectBox.setEnabled(useAiSupportBox.isSelected());
      for (JRadioButton rButton : radioButtons) {
        rButton.setEnabled(useAiSupportBox.isSelected() && autoCorrectBox.isSelected());
      }
      aiAutoSuggestionBox.setEnabled(useAiSupportBox.isSelected() && autoCorrectBox.isSelected());
    });
    
    aiUrlField.setEnabled(config.useAiSupport());
    modelField.setEnabled(config.useAiSupport());
    apiKeyField.setEnabled(config.useAiSupport());
    autoCorrectBox.setEnabled(config.useAiSupport());
    for (JRadioButton rButton : radioButtons) {
      rButton.setEnabled(config.useAiSupport() && config.aiAutoCorrect());
    }

    JLabel experimentalHint = new JLabel(messages.getString("guiAiExperimentalHint"));
    experimentalHint.setForeground(Color.red);
    cons.gridy++;
    aiOptionPanel.add(experimentalHint, cons);
    JLabel qualityHint = new JLabel(messages.getString("guiAiQualityHint"));
    if (config.getThemeSelection() == WtGeneralTools.THEME_FLATDARK) {
      qualityHint.setForeground(WtConfiguration.HINT_COLOR_BLUE);
    } else {
      qualityHint.setForeground(Color.blue);
    }
    cons.gridy++;
    aiOptionPanel.add(qualityHint, cons);
    cons.gridy++;
    aiOptionPanel.add(getAiInstallationHint(), cons);
    cons.insets = new Insets(16, SHIFT2, 0, 0);
    cons.gridy++;
    aiOptionPanel.add(useAiSupportBox, cons);
    JPanel serverPanel = new JPanel();
    serverPanel.setLayout(new GridBagLayout());
    GridBagConstraints cons1 = new GridBagConstraints();
    cons1.insets = new Insets(YSHIFT1, SHIFT2, 0, 0);
    cons1.gridx = 0;
    cons1.gridy = 0;
    cons1.anchor = GridBagConstraints.WEST;
    cons1.fill = GridBagConstraints.NONE;
    cons1.weightx = 0.0f;
    serverPanel.add(otherUrlLabel, cons1);
    cons1.gridy++;
    cons1.insets = new Insets(0, SHIFT2, 0, 0);
    serverPanel.add(aiUrlField, cons1);
    cons1.gridy++;
    cons1.insets = new Insets(YSHIFT1, SHIFT2, 0, 0);
    serverPanel.add(modelLabel, cons1);
    cons1.gridy++;
    cons1.insets = new Insets(0, SHIFT2, 0, 0);
    serverPanel.add(modelField, cons1);
    cons1.gridy++;
    cons1.insets = new Insets(YSHIFT1, SHIFT2, 0, 0);
    serverPanel.add(apiKeyLabel, cons1);
    cons1.gridy++;
    cons1.insets = new Insets(0, SHIFT2, 0, 0);
    serverPanel.add(apiKeyField, cons1);

    cons.insets = new Insets(0, SHIFT2, 0, 0);
    cons.gridx = 0;
    cons.gridy++;
    aiOptionPanel.add(serverPanel, cons);
    
    cons.gridy++;
    cons.insets = new Insets(16, SHIFT2, 0, 0);
    aiOptionPanel.add(autoCorrectBox, cons);
    
    cons.insets = new Insets(0, SHIFT3, 0, 0);
    
    for (JRadioButton rButton : radioButtons) {
      cons.gridy++;
      aiOptionPanel.add(rButton, cons);
    }
    
    cons.gridy++;
    aiOptionPanel.add(aiAutoSuggestionBox, cons);

    cons.insets = new Insets(16, SHIFT2, 0, 0);
    cons.gridy++;
    JLabel grammarErrorColor = new JLabel(messages.getString("guiAiGrammarErrorColor") + ":");
    aiOptionPanel.add(grammarErrorColor, cons);
    
    cons.insets = new Insets(0, SHIFT2, 0, 0);
    cons.gridy++;
    aiOptionPanel.add(getColorPanel(WtOfficeTools.AI_GRAMMAR_CATEGORY, WtOfficeTools.AI_GRAMMAR_HINT_RULE_ID), cons);

    cons.insets = new Insets(12, SHIFT2, 0, 0);
    cons.gridy++;
    JLabel stylisticErrorColor = new JLabel(messages.getString("guiAiStylisticErrorColor") + ":");
    aiOptionPanel.add(stylisticErrorColor, cons);
    
    cons.insets = new Insets(0, SHIFT2, 0, 0);
    cons.gridy++;
    aiOptionPanel.add(getColorPanel(WtOfficeTools.AI_STYLE_CATEGORY, WtOfficeTools.AI_GRAMMAR_OTHER_RULE_ID), cons);

    return aiOptionPanel;
  }
  
  private JPanel getOfficeAiImgElements() {
    JPanel aiOptionPanel = new JPanel();
    aiOptionPanel.setLayout(new GridBagLayout());
    GridBagConstraints cons = new GridBagConstraints();
    cons.insets = new Insets(6, 6, 6, 6);
    cons.gridx = 0;
    cons.gridy = 0;
    cons.anchor = GridBagConstraints.NORTHWEST;
    cons.fill = GridBagConstraints.HORIZONTAL;
    cons.weightx = 10.0f;
    cons.weighty = 0.0f;
    
    JLabel otherUrlLabel = new JLabel(messages.getString("guiAiUrl") + ":");

    JTextField aiUrlField = new JTextField(config.aiImgUrl() ==  null ? "" : config.aiImgUrl(), 25);
    aiUrlField.setMinimumSize(new Dimension(100, 25));
    aiUrlField.getDocument().addDocumentListener(new DocumentListener() {
      @Override
      public void insertUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void removeUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void changedUpdate(DocumentEvent e) {
        String serverName = aiUrlField.getText();
        serverName = serverName.trim();
        if(serverName.isEmpty()) {
          serverName = null;
        }
        if (config.isValidAiServerUrl(serverName)) {
          aiUrlField.setForeground(Color.BLACK);
          config.setAiImgUrl(serverName);;
        } else {
          aiUrlField.setForeground(Color.RED);
        }
      }
    });

    JLabel modelLabel = new JLabel(messages.getString("guiAiModel") + ":");

    JTextField modelField = new JTextField(config.aiImgModel() ==  null ? "" : config.aiImgModel(), 25);
    modelField.setMinimumSize(new Dimension(100, 25));
    modelField.getDocument().addDocumentListener(new DocumentListener() {
      @Override
      public void insertUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void removeUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void changedUpdate(DocumentEvent e) {
        String model = modelField.getText();
        model = model.trim();
        if(model.isEmpty()) {
          model = null;
        }
        if (model != null) {
          config.setAiImgModel(model);
        }
      }
    });

    JLabel apiKeyLabel = new JLabel(messages.getString("guiAiApiKey") + ":");

    JTextField apiKeyField = new JTextField(config.aiImgApiKey() ==  null ? "" : config.aiImgApiKey(), 25);
    apiKeyField.setMinimumSize(new Dimension(100, 25));
    apiKeyField.getDocument().addDocumentListener(new DocumentListener() {
      @Override
      public void insertUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void removeUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void changedUpdate(DocumentEvent e) {
        String apiKey = apiKeyField.getText();
        apiKey = apiKey.trim();
        if(apiKey.isEmpty()) {
          apiKey = null;
        }
        if (apiKey != null) {
          config.setAiImgApiKey(apiKey);
        }
      }
    });
    
    JCheckBox useAiSupportBox = new JCheckBox(messages.getString("guiUseAiSupport"));
    useAiSupportBox.setSelected(config.useAiImgSupport());
    useAiSupportBox.addItemListener(e -> {
      config.setUseAiImgSupport(useAiSupportBox.isSelected());
      aiUrlField.setEnabled(useAiSupportBox.isSelected());
      modelField.setEnabled(useAiSupportBox.isSelected());
      apiKeyField.setEnabled(useAiSupportBox.isSelected());
    });
    
    aiUrlField.setEnabled(config.useAiImgSupport());
    modelField.setEnabled(config.useAiImgSupport());
    apiKeyField.setEnabled(config.useAiImgSupport());

    JLabel experimentalHint = new JLabel(messages.getString("guiAiExperimentalHint"));
    experimentalHint.setForeground(Color.red);
    cons.gridy++;
    aiOptionPanel.add(experimentalHint, cons);
    JLabel qualityHint = new JLabel(messages.getString("guiAiQualityHint"));
    if (config.getThemeSelection() == WtGeneralTools.THEME_FLATDARK) {
      qualityHint.setForeground(WtConfiguration.HINT_COLOR_BLUE);
    } else {
      qualityHint.setForeground(Color.blue);
    }
    cons.gridy++;
    aiOptionPanel.add(qualityHint, cons);
    cons.gridy++;
    aiOptionPanel.add(getAiInstallationHint(), cons);
    cons.insets = new Insets(16, SHIFT2, 0, 0);
    cons.gridy++;
    aiOptionPanel.add(useAiSupportBox, cons);
    JPanel serverPanel = new JPanel();
    serverPanel.setLayout(new GridBagLayout());
    GridBagConstraints cons1 = new GridBagConstraints();
    cons1.insets = new Insets(YSHIFT1, SHIFT2, 0, 0);
    cons1.gridx = 0;
    cons1.gridy = 0;
    cons1.anchor = GridBagConstraints.WEST;
    cons1.fill = GridBagConstraints.NONE;
    cons1.weightx = 0.0f;
    serverPanel.add(otherUrlLabel, cons1);
    cons1.gridy++;
    cons1.insets = new Insets(0, SHIFT2, 0, 0);
    serverPanel.add(aiUrlField, cons1);
    cons1.gridy++;
    cons1.insets = new Insets(YSHIFT1, SHIFT2, 0, 0);
    serverPanel.add(modelLabel, cons1);
    cons1.gridy++;
    cons1.insets = new Insets(0, SHIFT2, 0, 0);
    serverPanel.add(modelField, cons1);
    cons1.gridy++;
    cons1.insets = new Insets(YSHIFT1, SHIFT2, 0, 0);
    serverPanel.add(apiKeyLabel, cons1);
    cons1.gridy++;
    cons1.insets = new Insets(0, SHIFT2, 0, 0);
    serverPanel.add(apiKeyField, cons1);

    cons.insets = new Insets(0, SHIFT2, 0, 0);
    cons.gridx = 0;
    cons.gridy++;
    aiOptionPanel.add(serverPanel, cons);
    cons.gridy++;
    cons.weighty = 10.0f;
    aiOptionPanel.add(new JLabel(" "), cons);
    
    return aiOptionPanel;
  }
  
  private JPanel getOfficeAiTtsElements() {
    JPanel aiOptionPanel = new JPanel();
    aiOptionPanel.setLayout(new GridBagLayout());
    GridBagConstraints cons = new GridBagConstraints();
    cons.insets = new Insets(6, 6, 6, 6);
    cons.gridx = 0;
    cons.gridy = 0;
    cons.anchor = GridBagConstraints.NORTHWEST;
    cons.fill = GridBagConstraints.HORIZONTAL;
    cons.weightx = 10.0f;
    cons.weighty = 0.0f;
    
    JLabel otherUrlLabel = new JLabel(messages.getString("guiAiUrl") + ":");

    JTextField aiUrlField = new JTextField(config.aiTtsUrl() ==  null ? "" : config.aiTtsUrl(), 25);
    aiUrlField.setMinimumSize(new Dimension(100, 25));
    aiUrlField.getDocument().addDocumentListener(new DocumentListener() {
      @Override
      public void insertUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void removeUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void changedUpdate(DocumentEvent e) {
        String serverName = aiUrlField.getText();
        serverName = serverName.trim();
        if(serverName.isEmpty()) {
          serverName = null;
        }
        if (config.isValidAiServerUrl(serverName)) {
          aiUrlField.setForeground(Color.BLACK);
          config.setAiTtsUrl(serverName);;
        } else {
          aiUrlField.setForeground(Color.RED);
        }
      }
    });

    JLabel modelLabel = new JLabel(messages.getString("guiAiModel") + ":");

    JTextField modelField = new JTextField(config.aiTtsModel() ==  null ? "" : config.aiTtsModel(), 25);
    modelField.setMinimumSize(new Dimension(100, 25));
    modelField.getDocument().addDocumentListener(new DocumentListener() {
      @Override
      public void insertUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void removeUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void changedUpdate(DocumentEvent e) {
        String model = modelField.getText();
        model = model.trim();
        if(model.isEmpty()) {
          model = null;
        }
        if (model != null) {
          config.setAiTtsModel(model);
        }
      }
    });

    JLabel apiKeyLabel = new JLabel(messages.getString("guiAiApiKey") + ":");

    JTextField apiKeyField = new JTextField(config.aiTtsApiKey() ==  null ? "" : config.aiTtsApiKey(), 25);
    apiKeyField.setMinimumSize(new Dimension(100, 25));
    apiKeyField.getDocument().addDocumentListener(new DocumentListener() {
      @Override
      public void insertUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void removeUpdate(DocumentEvent e) {
        changedUpdate(e);
      }
      @Override
      public void changedUpdate(DocumentEvent e) {
        String apiKey = apiKeyField.getText();
        apiKey = apiKey.trim();
        if(apiKey.isEmpty()) {
          apiKey = null;
        }
        if (apiKey != null) {
          config.setAiTtsApiKey(apiKey);
        }
      }
    });
    
    JCheckBox useAiSupportBox = new JCheckBox(messages.getString("guiUseAiSupport"));
    useAiSupportBox.setSelected(config.useAiTtsSupport());
    useAiSupportBox.addItemListener(e -> {
      config.setUseAiTtsSupport(useAiSupportBox.isSelected());
      aiUrlField.setEnabled(useAiSupportBox.isSelected());
      modelField.setEnabled(useAiSupportBox.isSelected());
      apiKeyField.setEnabled(useAiSupportBox.isSelected());
    });
    
    aiUrlField.setEnabled(config.useAiTtsSupport());
    modelField.setEnabled(config.useAiTtsSupport());
    apiKeyField.setEnabled(config.useAiTtsSupport());

    JLabel experimentalHint = new JLabel(messages.getString("guiAiExperimentalHint"));
    experimentalHint.setForeground(Color.red);
    cons.gridy++;
    aiOptionPanel.add(experimentalHint, cons);
    JLabel qualityHint = new JLabel(messages.getString("guiAiQualityHint"));
    if (config.getThemeSelection() == WtGeneralTools.THEME_FLATDARK) {
      qualityHint.setForeground(WtConfiguration.HINT_COLOR_BLUE);
    } else {
      qualityHint.setForeground(Color.blue);
    }
    cons.gridy++;
    aiOptionPanel.add(qualityHint, cons);
    cons.gridy++;
    aiOptionPanel.add(getAiInstallationHint(), cons);
    cons.insets = new Insets(16, SHIFT2, 0, 0);
    cons.gridy++;
    aiOptionPanel.add(useAiSupportBox, cons);
    JPanel serverPanel = new JPanel();
    serverPanel.setLayout(new GridBagLayout());
    GridBagConstraints cons1 = new GridBagConstraints();
    cons1.insets = new Insets(YSHIFT1, SHIFT2, 0, 0);
    cons1.gridx = 0;
    cons1.gridy = 0;
    cons1.anchor = GridBagConstraints.WEST;
    cons1.fill = GridBagConstraints.NONE;
    cons1.weightx = 0.0f;
    serverPanel.add(otherUrlLabel, cons1);
    cons1.gridy++;
    cons1.insets = new Insets(0, SHIFT2, 0, 0);
    serverPanel.add(aiUrlField, cons1);
    cons1.gridy++;
    cons1.insets = new Insets(YSHIFT1, SHIFT2, 0, 0);
    serverPanel.add(modelLabel, cons1);
    cons1.gridy++;
    cons1.insets = new Insets(0, SHIFT2, 0, 0);
    serverPanel.add(modelField, cons1);
    cons1.gridy++;
    cons1.insets = new Insets(YSHIFT1, SHIFT2, 0, 0);
    serverPanel.add(apiKeyLabel, cons1);
    cons1.gridy++;
    cons1.insets = new Insets(0, SHIFT2, 0, 0);
    serverPanel.add(apiKeyField, cons1);

    cons.insets = new Insets(0, SHIFT2, 0, 0);
    cons.gridx = 0;
    cons.gridy++;
    aiOptionPanel.add(serverPanel, cons);

    cons.gridy++;
    cons.weighty = 10.0f;
    aiOptionPanel.add(new JLabel(" "), cons);
    
    return aiOptionPanel;
  }
  
  private JTabbedPane getOfficeAiElements() {
    JTabbedPane tabbedpane = new JTabbedPane();
    tabbedpane.add(messages.getString("guiAiText"), getOfficeAiTextElements());
    tabbedpane.add(messages.getString("guiAiImages"), getOfficeAiImgElements());
    tabbedpane.add(messages.getString("guiAiTextToSpeech"), getOfficeAiTtsElements());
    return tabbedpane;
  }
  
}
