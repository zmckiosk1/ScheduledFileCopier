package scheduledfilecopier;

import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.text.*;
import java.util.*;
import java.util.List;
import java.util.Timer;
import java.util.TimerTask;
import java.util.stream.Collectors;

public class FileCopierGUI extends javax.swing.JFrame implements FileCopier.ProgressUpdater {

    private FileCopier fileCopier;
    private final DecimalFormat sizeFormat = new DecimalFormat("#,##0.00");
    private static final String CONFIG_FILE = "filecopier_settings.properties";
    private Timer scheduleTimer;
    private Date scheduledTime;
    private boolean dailySchedule = false;
    private JCheckBox dailyCheckbox;

    // GUI Components
    private javax.swing.JButton browseDest;
    private javax.swing.JButton browseSource;
    private javax.swing.JLabel bytesLabel;
    private javax.swing.JLabel currentFileLabel;
    private javax.swing.JTextField destField;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JScrollPane jScrollPane3;
    private javax.swing.JCheckBox lockedCheckbox;
    private javax.swing.JCheckBox forceCloseCheckbox;
    private javax.swing.JCheckBox vssCheckbox;
    private javax.swing.JTextArea logArea;
    private javax.swing.JProgressBar progressBar;
    private javax.swing.JLabel progressLabel;
    private javax.swing.JButton saveButton;
    private javax.swing.JTextField sourceField;
    private javax.swing.JButton startButton;
    private javax.swing.JButton stopButton;
    private javax.swing.JTextArea skipLocationsArea;
    private javax.swing.JTextArea priorityItemsArea;
    private javax.swing.JButton scheduleButton;
    private javax.swing.JSpinner timeSpinner;
    private javax.swing.JButton cancelScheduleButton;
    private javax.swing.JLabel nextRunLabel;
    private javax.swing.JButton rescheduleButton;

    public FileCopierGUI() {
        initComponents();
        loadSettings();
    }

    private void initComponents() {
        jLabel1 = new javax.swing.JLabel("Source File/Folder:");
        sourceField = new javax.swing.JTextField(30);
        jLabel2 = new javax.swing.JLabel("Destination:");
        destField = new javax.swing.JTextField(30);
        lockedCheckbox = new javax.swing.JCheckBox("Copy locked files (e.g., .pst)");
        forceCloseCheckbox = new javax.swing.JCheckBox("Force close applications locking files");
        vssCheckbox = new javax.swing.JCheckBox("Use Volume Shadow Copy (Admin required)");
        startButton = new javax.swing.JButton("Start Copy");
        stopButton = new javax.swing.JButton("Stop");
        browseSource = new javax.swing.JButton("Browse...");
        browseDest = new javax.swing.JButton("Browse...");
        progressBar = new javax.swing.JProgressBar(0, 100);
        currentFileLabel = new javax.swing.JLabel(" ");
        progressLabel = new javax.swing.JLabel(" ");
        bytesLabel = new javax.swing.JLabel(" ");
        logArea = new javax.swing.JTextArea(10, 40);
        jScrollPane1 = new javax.swing.JScrollPane(logArea);
        jLabel3 = new javax.swing.JLabel("Priority files/folders (one per line, copied first):");
        priorityItemsArea = new javax.swing.JTextArea(5, 40);
        jScrollPane3 = new javax.swing.JScrollPane(priorityItemsArea);
        jLabel4 = new javax.swing.JLabel("Skip locations (one per line, partial matches OK):");
        skipLocationsArea = new javax.swing.JTextArea(5, 40);
        jScrollPane2 = new javax.swing.JScrollPane(skipLocationsArea);
        saveButton = new javax.swing.JButton("Save Settings");
        jLabel5 = new javax.swing.JLabel("Schedule Time:");
        timeSpinner = new JSpinner(new SpinnerDateModel());
        JSpinner.DateEditor timeEditor = new JSpinner.DateEditor(timeSpinner, "HH:mm");
        timeSpinner.setEditor(timeEditor);
        timeSpinner.setValue(new Date());
        scheduleButton = new javax.swing.JButton("Schedule Copy");
        cancelScheduleButton = new javax.swing.JButton("Cancel Schedule");
        rescheduleButton = new javax.swing.JButton("Reschedule");
        cancelScheduleButton.setEnabled(false);
        rescheduleButton.setEnabled(false);
        nextRunLabel = new javax.swing.JLabel("Next scheduled run: Not scheduled");
        dailyCheckbox = new JCheckBox("Repeat daily");

        logArea.setEditable(false);
        stopButton.setEnabled(false);
        forceCloseCheckbox.setEnabled(false);
        priorityItemsArea.setLineWrap(true);
        skipLocationsArea.setLineWrap(true);
        
        boolean isWindows = System.getProperty("os.name").toLowerCase().contains("win");
        vssCheckbox.setVisible(isWindows);
        
        lockedCheckbox.addActionListener(e -> {
            forceCloseCheckbox.setEnabled(lockedCheckbox.isSelected());
            vssCheckbox.setEnabled(lockedCheckbox.isSelected() && isWindows);
            if (!lockedCheckbox.isSelected()) {
                forceCloseCheckbox.setSelected(false);
                vssCheckbox.setSelected(false);
            }
        });

        startButton.addActionListener(e -> startButtonActionPerformed());
        stopButton.addActionListener(e -> stopButtonActionPerformed());
        browseSource.addActionListener(e -> browseSourceActionPerformed());
        browseDest.addActionListener(e -> browseDestActionPerformed());
        saveButton.addActionListener(e -> saveButtonActionPerformed());
        scheduleButton.addActionListener(e -> scheduleButtonActionPerformed());
        cancelScheduleButton.addActionListener(e -> cancelScheduleButtonActionPerformed());
        rescheduleButton.addActionListener(e -> rescheduleButtonActionPerformed());

        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.anchor = GridBagConstraints.WEST;
        gbc.fill = GridBagConstraints.HORIZONTAL;

        gbc.gridx = 0; gbc.gridy = 0;
        panel.add(jLabel1, gbc);
        gbc.gridx = 1; gbc.weightx = 1;
        panel.add(sourceField, gbc);
        gbc.gridx = 2; gbc.weightx = 0;
        panel.add(browseSource, gbc);

        gbc.gridx = 0; gbc.gridy = 1;
        panel.add(jLabel2, gbc);
        gbc.gridx = 1;
        panel.add(destField, gbc);
        gbc.gridx = 2;
        panel.add(browseDest, gbc);

        gbc.gridx = 0; gbc.gridy = 2; gbc.gridwidth = 3;
        panel.add(lockedCheckbox, gbc);

        gbc.gridy = 3;
        panel.add(forceCloseCheckbox, gbc);

        gbc.gridy = 4;
        panel.add(vssCheckbox, gbc);

        gbc.gridy = 5;
        panel.add(jLabel3, gbc);

        gbc.gridy = 6; gbc.fill = GridBagConstraints.BOTH; gbc.weighty = 1;
        panel.add(jScrollPane3, gbc);
        gbc.fill = GridBagConstraints.HORIZONTAL; gbc.weighty = 0;

        gbc.gridy = 7;
        panel.add(jLabel4, gbc);

        gbc.gridy = 8; gbc.fill = GridBagConstraints.BOTH; gbc.weighty = 1;
        panel.add(jScrollPane2, gbc);
        gbc.fill = GridBagConstraints.HORIZONTAL; gbc.weighty = 0;

        gbc.gridy = 9;
        panel.add(new JSeparator(), gbc);

        gbc.gridy = 10;
        JPanel schedulePanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        schedulePanel.add(jLabel5);
        schedulePanel.add(timeSpinner);
        schedulePanel.add(scheduleButton);
        schedulePanel.add(cancelScheduleButton);
        schedulePanel.add(rescheduleButton);
        schedulePanel.add(dailyCheckbox);
        panel.add(schedulePanel, gbc);
        
        gbc.gridy = 11;
        panel.add(nextRunLabel, gbc);
        
        gbc.gridy = 12;
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        buttonPanel.add(startButton);
        buttonPanel.add(stopButton);
        buttonPanel.add(saveButton);
        panel.add(buttonPanel, gbc);

        gbc.gridy = 13;
        panel.add(currentFileLabel, gbc);

        gbc.gridy = 14;
        panel.add(progressBar, gbc);

        gbc.gridy = 15;
        JPanel progressPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        progressPanel.add(progressLabel);
        progressPanel.add(bytesLabel);
        panel.add(progressPanel, gbc);

        gbc.gridy = 16; gbc.fill = GridBagConstraints.BOTH; gbc.weighty = 1;
        panel.add(jScrollPane1, gbc);

        add(panel);
        pack();
        setLocationRelativeTo(null);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setTitle("Scheduled File Copier with Priority Items");
    }

    private void scheduleButtonActionPerformed() {
        Date selectedTime = (Date) timeSpinner.getValue();
        scheduledTime = calculateNextRunTime(selectedTime);
        dailySchedule = dailyCheckbox.isSelected();
        
        scheduleTimer = new Timer(true);
        long delay = scheduledTime.getTime() - System.currentTimeMillis();
        
        scheduleTimer.schedule(createScheduledTask(), delay);
        
        scheduleButton.setEnabled(false);
        cancelScheduleButton.setEnabled(true);
        rescheduleButton.setEnabled(true);
        dailyCheckbox.setEnabled(false);
        nextRunLabel.setText("Next scheduled run: " + scheduledTime);
        log("Copy scheduled for " + scheduledTime + (dailySchedule ? " (Daily)" : ""));
        
        saveSettings();
    }

    private Date calculateNextRunTime(Date selectedTime) {
        Calendar cal = Calendar.getInstance();
        Calendar selectedCal = Calendar.getInstance();
        selectedCal.setTime(selectedTime);
        
        cal.set(Calendar.HOUR_OF_DAY, selectedCal.get(Calendar.HOUR_OF_DAY));
        cal.set(Calendar.MINUTE, selectedCal.get(Calendar.MINUTE));
        cal.set(Calendar.SECOND, 0);
        cal.set(Calendar.MILLISECOND, 0);
        
        if (cal.getTime().before(new Date())) {
            cal.add(Calendar.DAY_OF_MONTH, 1);
        }
        
        return cal.getTime();
    }

    private TimerTask createScheduledTask() {
        return new TimerTask() {
            @Override
            public void run() {
                SwingUtilities.invokeLater(() -> {
                    log("Scheduled copy started at " + new Date());
                    startButtonActionPerformed();
                });
                
                if (dailySchedule && scheduleTimer != null) {
                    Calendar nextCal = Calendar.getInstance();
                    nextCal.setTime(scheduledTime);
                    nextCal.add(Calendar.DAY_OF_MONTH, 1);
                    scheduledTime = nextCal.getTime();
                    long nextDelay = scheduledTime.getTime() - System.currentTimeMillis();
                    
                    scheduleTimer.schedule(createScheduledTask(), nextDelay);
                    
                    SwingUtilities.invokeLater(() -> {
                        nextRunLabel.setText("Next scheduled run: " + scheduledTime);
                        saveSettings();
                    });
                }
            }
        };
    }

    private void rescheduleButtonActionPerformed() {
        if (scheduledTime != null) {
            cancelScheduleButtonActionPerformed();
            timeSpinner.setValue(scheduledTime);
            scheduleButtonActionPerformed();
        }
    }

    private void cancelScheduleButtonActionPerformed() {
        if (scheduleTimer != null) {
            scheduleTimer.cancel();
            scheduleTimer = null;
        }
        scheduledTime = null;
        scheduleButton.setEnabled(true);
        cancelScheduleButton.setEnabled(false);
        rescheduleButton.setEnabled(false);
        dailyCheckbox.setEnabled(true);
        nextRunLabel.setText("Next scheduled run: Not scheduled");
        log("Scheduled copy cancelled");
        
        saveSettings();
    }

    private void startButtonActionPerformed() {
        String source = sourceField.getText();
        String dest = destField.getText();
        boolean copyLocked = lockedCheckbox.isSelected();
        boolean forceClose = forceCloseCheckbox.isSelected();
        boolean useVSS = vssCheckbox.isSelected();
        
        if (source.isEmpty() || dest.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please specify source and destination paths");
            return;
        }
        
        if (source.toLowerCase().endsWith(".pst") || (new File(source).isDirectory() && containsPstFiles(new File(source)))) {
            int result = JOptionPane.showConfirmDialog(this,
                "<html><b>Outlook PST File Detected</b><br><br>" +
                "For successful copying:<ol>" +
                "<li>Close Outlook completely (check Task Manager)</li>" +
                "<li>Run this program as Administrator</li>" +
                "<li>Enable 'Use Volume Shadow Copy'</li>" +
                "<li>Disable Outlook Add-ins if issues persist</li></ol>" +
                "Continue with copy operation?</html>",
                "PST File Warning",
                JOptionPane.YES_NO_OPTION,
                JOptionPane.WARNING_MESSAGE);
            
            if (result != JOptionPane.YES_OPTION) {
                return;
            }
        }
        
        final List<String> priorityItems = Arrays.stream(
                Optional.ofNullable(priorityItemsArea.getText()).orElse("").split("\\r?\\n"))
            .map(String::trim)
            .filter(line -> !line.isEmpty())
            .collect(Collectors.toList());
        
        final List<String> skipLocations = Arrays.stream(
                Optional.ofNullable(skipLocationsArea.getText()).orElse("").split("\\r?\\n"))
            .map(String::trim)
            .filter(line -> !line.isEmpty())
            .collect(Collectors.toList());
        
        SwingUtilities.invokeLater(() -> {
            progressBar.setValue(0);
            currentFileLabel.setText("Preparing to copy...");
            progressLabel.setText("0%");
            bytesLabel.setText("0 B / 0 B");
            log("Starting copy operation...");
            
            startButton.setEnabled(false);
            stopButton.setEnabled(true);
            saveButton.setEnabled(false);
            scheduleButton.setEnabled(false);
        });
        
        new Thread(() -> {
            try {
                fileCopier = new FileCopier(source, dest, copyLocked, forceClose, useVSS, 
                                          priorityItems, skipLocations, this);
                fileCopier.startCopy();
                log("Copy completed successfully!");
            } catch (Exception ex) {
                logError("Error during copy: " + ex.getMessage());
                
                if (ex.getMessage().contains(".pst")) {
                    logError("<html><b>PST COPY FAILED</b><br>" +
                        "1. Completely close Outlook (check Task Manager)<br>" +
                        "2. Run as Administrator<br>" +
                        "3. Enable Volume Shadow Copy<br>" +
                        "4. Try copying individual PST files</html>");
                }
            } finally {
                SwingUtilities.invokeLater(() -> {
                    startButton.setEnabled(true);
                    stopButton.setEnabled(false);
                    saveButton.setEnabled(true);
                    if (scheduleTimer == null) {
                        scheduleButton.setEnabled(true);
                    }
                });
            }
        }).start();
    }

    private boolean containsPstFiles(File directory) {
        if (!directory.isDirectory()) return false;
        
        File[] files = directory.listFiles();
        if (files != null) {
            for (File file : files) {
                if (file.isDirectory()) {
                    if (containsPstFiles(file)) return true;
                } else if (file.getName().toLowerCase().endsWith(".pst")) {
                    return true;
                }
            }
        }
        return false;
    }

    private void stopButtonActionPerformed() {
        if (fileCopier != null) {
            fileCopier.cancelCopy();
            stopButton.setEnabled(false);
            log("Copy operation cancelled by user");
        }
    }

    private void saveSettings() {
        Properties props = new Properties();
        props.setProperty("source", sourceField.getText());
        props.setProperty("destination", destField.getText());
        props.setProperty("copyLocked", Boolean.toString(lockedCheckbox.isSelected()));
        props.setProperty("forceClose", Boolean.toString(forceCloseCheckbox.isSelected()));
        props.setProperty("useVSS", Boolean.toString(vssCheckbox.isSelected()));
        props.setProperty("priorityItems", priorityItemsArea.getText().replace("\n", "|||"));
        props.setProperty("skipLocations", skipLocationsArea.getText().replace("\n", "|||"));
        props.setProperty("dailySchedule", Boolean.toString(dailyCheckbox.isSelected()));
        
        if (scheduledTime != null) {
            props.setProperty("scheduledTime", Long.toString(scheduledTime.getTime()));
            props.setProperty("scheduledDaily", Boolean.toString(dailySchedule));
        } else {
            props.remove("scheduledTime");
            props.remove("scheduledDaily");
        }
        
        try (OutputStream output = new FileOutputStream(CONFIG_FILE)) {
            props.store(output, "File Copier Settings");
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this, "Error saving settings: " + ex.getMessage());
        }
    }

    private void loadSettings() {
        File configFile = new File(CONFIG_FILE);
        if (!configFile.exists()) {
            return;
        }

        Properties props = new Properties();
        try (InputStream input = new FileInputStream(CONFIG_FILE)) {
            props.load(input);

            sourceField.setText(props.getProperty("source", ""));
            destField.setText(props.getProperty("destination", ""));
            lockedCheckbox.setSelected(Boolean.parseBoolean(props.getProperty("copyLocked", "false")));
            forceCloseCheckbox.setSelected(Boolean.parseBoolean(props.getProperty("forceClose", "false")));
            vssCheckbox.setSelected(Boolean.parseBoolean(props.getProperty("useVSS", "false")));
            dailyCheckbox.setSelected(Boolean.parseBoolean(props.getProperty("dailySchedule", "false")));
            
            boolean isWindows = System.getProperty("os.name").toLowerCase().contains("win");
            forceCloseCheckbox.setEnabled(lockedCheckbox.isSelected());
            vssCheckbox.setEnabled(lockedCheckbox.isSelected() && isWindows);

            String priorityItems = props.getProperty("priorityItems", "").replace("|||", "\n");
            priorityItemsArea.setText(priorityItems);
            
            String skipLocations = props.getProperty("skipLocations", "").replace("|||", "\n");
            skipLocationsArea.setText(skipLocations);
            
            String scheduledTimeStr = props.getProperty("scheduledTime");
            if (scheduledTimeStr != null && !scheduledTimeStr.isEmpty()) {
                long time = Long.parseLong(scheduledTimeStr);
                scheduledTime = new Date(time);
                dailySchedule = Boolean.parseBoolean(props.getProperty("scheduledDaily", "false"));
                dailyCheckbox.setSelected(dailySchedule);
                
                timeSpinner.setValue(scheduledTime);
                
                if (scheduledTime.before(new Date())) {
                    Calendar cal = Calendar.getInstance();
                    cal.setTime(scheduledTime);
                    cal.add(Calendar.DAY_OF_MONTH, 1);
                    scheduledTime = cal.getTime();
                }
                
                scheduleButton.setEnabled(true);
                cancelScheduleButton.setEnabled(false);
                rescheduleButton.setEnabled(true);
                nextRunLabel.setText("Next scheduled run: " + scheduledTime);
            }
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this, "Error loading settings: " + ex.getMessage());
        } catch (NumberFormatException ex) {
            log("Error parsing scheduled time from settings");
        }
    }

    @Override
    public void updateProgress(String currentFile, int progress, long bytesCopied, long totalBytes) {
        SwingUtilities.invokeLater(() -> {
            currentFileLabel.setText(currentFile);
            progressBar.setValue(progress);
            progressLabel.setText(progress + "%");
            
            String bytesText = formatSize(bytesCopied) + " / " + formatSize(totalBytes);
            bytesLabel.setText(bytesText);
        });
    }

    @Override
    public void logMessage(String message) {
        SwingUtilities.invokeLater(() -> {
            logArea.append(message + "\n");
            logArea.setCaretPosition(logArea.getDocument().getLength());
        });
    }

    private void log(String message) {
        logMessage(new SimpleDateFormat("[yyyy-MM-dd HH:mm:ss] ").format(new Date()) + message);
    }

    private void logError(String message) {
        SwingUtilities.invokeLater(() -> {
            logArea.append(new SimpleDateFormat("[yyyy-MM-dd HH:mm:ss] ERROR: ").format(new Date()) + message + "\n");
            logArea.setCaretPosition(logArea.getDocument().getLength());
        });
    }

    @Override
    public boolean isCancelled() {
        return stopButton.isEnabled() && !startButton.isEnabled();
    }

    private String formatSize(long bytes) {
        if (bytes < 1024) return bytes + " B";
        if (bytes < 1024 * 1024) return sizeFormat.format(bytes / 1024.0) + " KB";
        if (bytes < 1024 * 1024 * 1024) return sizeFormat.format(bytes / (1024.0 * 1024)) + " MB";
        return sizeFormat.format(bytes / (1024.0 * 1024 * 1024)) + " GB";
    }

    private void browseSourceActionPerformed() {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            sourceField.setText(chooser.getSelectedFile().getAbsolutePath());
        }
    }

    private void browseDestActionPerformed() {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            destField.setText(chooser.getSelectedFile().getAbsolutePath());
        }
    }

    private void saveButtonActionPerformed() {
        saveSettings();
        JOptionPane.showMessageDialog(this, "Settings saved successfully!");
    }

    public static void main(String[] args) {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        
        SwingUtilities.invokeLater(() -> {
            FileCopierGUI gui = new FileCopierGUI();
            gui.setVisible(true);
        });
    }
}
