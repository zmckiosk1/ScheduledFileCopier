package scheduledfilecopier;

import java.io.*;
import java.nio.channels.FileChannel;
import java.nio.file.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import javax.swing.SwingUtilities;

public class FileCopier {

    private String sourcePath;
    private String destinationPath;
    private boolean copyLockedFiles;
    private boolean forceClose;
    private boolean useVSS;
    private ProgressUpdater progressUpdater;
    private long totalBytesToCopy;
    private long bytesCopied;
    private boolean isCancelled;
    private List<String> skipLocations;
    private List<String> priorityItems;

    public interface ProgressUpdater {
        void updateProgress(String currentFile, int progress, long bytesCopied, long totalBytes);
        void logMessage(String message);
        boolean isCancelled();
    }

    public FileCopier(String sourcePath, String destinationPath, boolean copyLockedFiles, 
                     boolean forceClose, boolean useVSS, 
                     List<String> priorityItems, List<String> skipLocations, 
                     ProgressUpdater progressUpdater) {
        this.sourcePath = sourcePath;
        this.destinationPath = destinationPath;
        this.copyLockedFiles = copyLockedFiles;
        this.forceClose = forceClose;
        this.useVSS = useVSS && System.getProperty("os.name").toLowerCase().contains("win");
        this.progressUpdater = progressUpdater;
        this.priorityItems = priorityItems != null ? priorityItems : new ArrayList<>();
        this.skipLocations = skipLocations != null ? skipLocations : new ArrayList<>();
    }

    public void startCopy() throws IOException {
        isCancelled = false;
        bytesCopied = 0;
        calculateTotalBytes();
        
        // First copy priority items
        copyPriorityItems();
        
        // Then copy the main source
        File source = new File(sourcePath);
        File dest = new File(destinationPath);
        
        if (!source.exists()) {
            throw new IOException("Source path does not exist");
        }
        
        if (source.isFile()) {
            copySingleFile(source, dest);
        } else {
            copyDirectory(source, dest);
        }
    }

    private void copyPriorityItems() throws IOException {
        for (String priorityItem : priorityItems) {
            if (isCancelled) break;
            
            File sourceFile = new File(priorityItem);
            if (!sourceFile.exists()) {
                log("Priority item not found: " + priorityItem);
                continue;
            }

            if (shouldSkip(sourceFile)) {
                log("Skipping priority item (matches skip pattern): " + priorityItem);
                continue;
            }

            // Calculate destination path
            String relativePath;
            if (priorityItem.startsWith(sourcePath)) {
                relativePath = sourcePath.equals(priorityItem) 
                    ? "" 
                    : priorityItem.substring(sourcePath.length() + 1);
            } else {
                relativePath = "Priority Items" + File.separator + sourceFile.getName();
            }

            File destFile = new File(destinationPath, relativePath);
            
            if (sourceFile.isDirectory()) {
                // Create parent directories if they don't exist
                if (!destFile.exists() && !destFile.mkdirs()) {
                    throw new IOException("Failed to create directory: " + destFile.getAbsolutePath());
                }
                
                // Copy directory contents recursively
                copyDirectoryContents(sourceFile, destFile);
            } else {
                // Ensure parent directory exists
                File parent = destFile.getParentFile();
                if (parent != null && !parent.exists() && !parent.mkdirs()) {
                    throw new IOException("Failed to create directory: " + parent.getAbsolutePath());
                }
                
                copySingleFile(sourceFile, destFile);
            }
        }
    }

    private void copyDirectoryContents(File sourceDir, File destDir) throws IOException {
        File[] files = sourceDir.listFiles();
        if (files == null) return;

        for (File file : files) {
            if (isCancelled) return;
            if (shouldSkip(file)) continue;
            
            File destFile = new File(destDir, file.getName());
            if (file.isDirectory()) {
                copyDirectoryContents(file, destFile);
            } else {
                copySingleFile(file, destFile);
            }
        }
    }

    public void cancelCopy() {
        isCancelled = true;
    }

    private void calculateTotalBytes() {
        File source = new File(sourcePath);
        totalBytesToCopy = 0;
        
        // Calculate size for priority items first
        for (String priorityItem : priorityItems) {
            File file = new File(priorityItem);
            if (file.exists() && !shouldSkip(file)) {
                if (file.isFile()) {
                    totalBytesToCopy += file.length();
                } else {
                    totalBytesToCopy += calculateDirectorySize(file);
                }
            }
        }
        
        // Then calculate size for main source (excluding any priority items that are within it)
        if (source.isFile()) {
            if (!shouldSkip(source) && !isPriorityItem(source.getAbsolutePath())) {
                totalBytesToCopy += source.length();
            }
        } else {
            totalBytesToCopy += calculateDirectorySizeExcludingPriorityItems(source);
        }
        
        log("Total bytes to copy: " + totalBytesToCopy);
    }

    private long calculateDirectorySizeExcludingPriorityItems(File directory) {
        long size = 0;
        File[] files = directory.listFiles();
        if (files != null) {
            for (File file : files) {
                if (isCancelled) return 0;
                if (shouldSkip(file) || isPriorityItem(file.getAbsolutePath())) continue;
                
                if (file.isFile()) {
                    size += file.length();
                } else {
                    size += calculateDirectorySizeExcludingPriorityItems(file);
                }
            }
        }
        return size;
    }

    private boolean isPriorityItem(String path) {
        for (String priorityItem : priorityItems) {
            if (path.startsWith(priorityItem)) {
                return true;
            }
        }
        return false;
    }

    private long calculateDirectorySize(File directory) {
        long size = 0;
        File[] files = directory.listFiles();
        if (files != null) {
            for (File file : files) {
                if (isCancelled) return 0;
                if (shouldSkip(file)) continue;
                
                if (file.isFile()) {
                    size += file.length();
                } else {
                    size += calculateDirectorySize(file);
                }
            }
        }
        return size;
    }

    private boolean shouldSkip(File file) {
        String absolutePath = file.getAbsolutePath().toLowerCase();
        for (String skipLocation : skipLocations) {
            if (!skipLocation.trim().isEmpty() && absolutePath.contains(skipLocation.toLowerCase().trim())) {
                return true;
            }
        }
        return false;
    }

    private void copySingleFile(File source, File dest) throws IOException {
        if (isCancelled || shouldSkip(source)) return;
        
        if (dest.isDirectory()) {
            dest = new File(dest, source.getName());
        }
        
        updateProgress("Copying: " + source.getName(), 0);
        
        if (source.getName().toLowerCase().endsWith(".pst")) {
            copyPstFile(source, dest);
            return;
        }
        
        if (!copyLockedFiles && isFileLocked(source)) {
            throw new IOException("File is locked and copyLockedFiles is false: " + source.getAbsolutePath());
        }
        
        try {
            robustCopy(source, dest);
        } catch (IOException e) {
            if (copyLockedFiles && forceClose) {
                copyLockedFileWithForceClose(source, dest);
            } else if (copyLockedFiles) {
                copyLockedFile(source, dest);
            } else {
                throw e;
            }
        }
    }

    private void copyPstFile(File source, File dest) throws IOException {
        log("=== Starting PST Copy: " + source.getName() + " ===");
        log("Size: " + formatSize(source.length()));
        log("Last Modified: " + new Date(source.lastModified()));
        
        // 1. First try VSS if enabled and available
        if (useVSS && hasAdminPrivileges()) {
            try {
                log("Attempting Volume Shadow Copy (Recommended for PST files)...");
                copyWithVSS(source, dest);
                log("Successfully copied via VSS");
                return;
            } catch (IOException e) {
                log("VSS copy failed: " + e.getMessage());
                if (e.getMessage().contains("Access is denied")) {
                    log("Admin rights may be insufficient. Try running as Administrator.");
                }
            }
        }
        
        // 2. Verify Outlook is not running
        if (isOutlookRunning()) {
            log("Outlook appears to be running - attempting to close...");
            if (!enhancedCloseOutlookProcesses()) {
                throw new IOException("Could not close Outlook processes");
            }
            // Extra wait time after closing Outlook
            try { Thread.sleep(5000); } catch (InterruptedException ie) {}
        }
        
        // 3. Try normal copy with retries
        int retries = 3;
        for (int i = 1; i <= retries; i++) {
            try {
                log("Attempting normal copy (Attempt " + i + "/" + retries + ")...");
                copyWithStreams(source, dest);
                log("Successfully copied via normal method");
                return;
            } catch (IOException e) {
                log("Attempt " + i + " failed: " + e.getMessage());
                if (i < retries) {
                    try { Thread.sleep(2000); } catch (InterruptedException ie) {}
                }
            }
        }
        
        // 4. Final fallback with detailed error reporting
        try {
            log("Attempting locked file copy with read-only access...");
            copyLockedFile(source, dest);
            log("Successfully copied via locked file method");
        } catch (IOException e) {
            String errorReport = generatePstErrorReport(source);
            throw new IOException("Failed to copy PST file after all attempts.\n" + errorReport, e);
        }
    }

    private boolean isOutlookRunning() {
        try {
            Process process = Runtime.getRuntime().exec("tasklist /FI \"IMAGENAME eq outlook.exe\"");
            String output = readProcessOutput(process);
            return output.contains("outlook.exe");
        } catch (Exception e) {
            log("Error checking Outlook status: " + e.getMessage());
            return false;
        }
    }

    private boolean enhancedCloseOutlookProcesses() {
        try {
            log("Starting enhanced Outlook process termination...");
            
            // 1. First try graceful close
            log("Attempting to close Outlook gracefully...");
            Process graceful = Runtime.getRuntime().exec("taskkill /IM outlook.exe");
            graceful.waitFor();
            Thread.sleep(3000);
            
            // 2. Verify closure
            if (!isOutlookRunning()) {
                log("Outlook closed gracefully");
                return true;
            }
            
            // 3. Force kill if still running
            log("Forcing Outlook to close...");
            Process force = Runtime.getRuntime().exec("taskkill /IM outlook.exe /F");
            int exitCode = force.waitFor();
            Thread.sleep(5000); // Extended wait time
            
            // 4. Final verification
            if (!isOutlookRunning()) {
                log("Successfully forced Outlook to close");
                return true;
            }
            
            log("Failed to close Outlook completely (exit code: " + exitCode + ")");
            return false;
        } catch (Exception e) {
            log("Error closing Outlook: " + e.getMessage());
            return false;
        }
    }

    private String generatePstErrorReport(File pstFile) {
        StringBuilder report = new StringBuilder();
        report.append("\n=== PST COPY FAILURE REPORT ===\n");
        report.append("File: ").append(pstFile.getAbsolutePath()).append("\n");
        report.append("Size: ").append(formatSize(pstFile.length())).append("\n");
        report.append("Modified: ").append(new Date(pstFile.lastModified())).append("\n\n");
        
        report.append("SYSTEM STATUS:\n");
        report.append("- OS: ").append(System.getProperty("os.name")).append("\n");
        report.append("- Java: ").append(System.getProperty("java.version")).append("\n");
        report.append("- Admin: ").append(hasAdminPrivileges() ? "Yes" : "No").append("\n");
        report.append("- VSS: ").append(useVSS ? "Enabled" : "Disabled").append("\n");
        report.append("- Outlook Running: ").append(isOutlookRunning() ? "Yes" : "No").append("\n\n");
        
        report.append("RECOMMENDED ACTIONS:\n");
        report.append("1. MANUALLY CLOSE OUTLOOK (check Task Manager)\n");
        report.append("2. Run this program as Administrator\n");
        report.append("3. Enable Volume Shadow Copy in settings\n");
        report.append("4. Try copying during non-business hours\n");
        report.append("5. Use Outlook's built-in export function\n");
        
        return report.toString();
    }

    private void copyWithVSS(File source, File dest) throws IOException {
        if (!System.getProperty("os.name").toLowerCase().contains("win")) {
            throw new IOException("VSS is only supported on Windows");
        }

        try {
            String sourcePath = source.getAbsolutePath().replace("'", "''");
            String destPath = dest.getAbsolutePath().replace("'", "''");
            
            // Enhanced VSS command with better error handling
            String cmd = String.format(
                "powershell -command \"$ErrorActionPreference='Stop'; " +
                "Try { " +
                "  $file = New-Object System.IO.FileInfo('%s'); " +
                "  $stream = $file.Open([System.IO.FileMode]::Open, " +
                "    [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite); " +
                "  Copy-Item -Path '%s' -Destination '%s' -Force; " +
                "  $stream.Close(); " +
                "  if (!(Test-Path '%s')) { exit 1 } " +
                "} Catch { exit 2 }\"",
                sourcePath, sourcePath, destPath, destPath);

            log("Executing VSS command...");
            Process process = Runtime.getRuntime().exec(cmd);
            
            // Read error stream
            StringBuilder errors = new StringBuilder();
            try (BufferedReader reader = new BufferedReader(
                new InputStreamReader(process.getErrorStream()))) {
                String line;
                while ((line = reader.readLine()) != null) {
                    errors.append(line).append("\n");
                }
            }
            
            int exitCode = process.waitFor();
            if (exitCode == 1) {
                throw new IOException("VSS copy failed - destination file not created");
            } else if (exitCode == 2) {
                throw new IOException("VSS copy failed with PowerShell error: " + errors.toString());
            } else if (exitCode != 0) {
                throw new IOException("VSS copy failed with unknown error (code " + exitCode + ")");
            }
            
            // Verify copy
            if (!dest.exists() || dest.length() != source.length()) {
                throw new IOException("VSS copy verification failed - file sizes don't match");
            }
            
            // Update progress
            bytesCopied += source.length();
            int progress = (int) ((bytesCopied * 100) / totalBytesToCopy);
            updateProgress("Copied (VSS): " + source.getName(), progress);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new IOException("VSS copy interrupted", e);
        }
    }

    private boolean hasAdminPrivileges() {
        if (!System.getProperty("os.name").toLowerCase().contains("win")) {
            return false;
        }
        
        try {
            String cmd = "net session >nul 2>&1";
            Process process = Runtime.getRuntime().exec(cmd);
            return process.waitFor() == 0;
        } catch (Exception e) {
            return false;
        }
    }

    private void robustCopy(File source, File dest) throws IOException {
        // 1. First try normal copy
        try {
            if (!copyLockedFiles) {
                copyWithFileChannels(source, dest);
            } else {
                copyWithStreams(source, dest);
            }
            return;
        } catch (IOException e) {
            if (!copyLockedFiles) throw e;
            log("Normal copy failed, attempting alternative approaches...");
        }
        
        // 2. Try Windows VSS if available and enabled
        if (useVSS && hasAdminPrivileges()) {
            try {
                copyWithVSS(source, dest);
                return;
            } catch (IOException e) {
                log("VSS copy failed: " + e.getMessage());
            }
        }
        
        // 3. Fall back to other methods
        if (forceClose) {
            copyLockedFileWithForceClose(source, dest);
        } else {
            copyLockedFile(source, dest);
        }
    }

    private void copyWithFileChannels(File source, File dest) throws IOException {
        try (FileChannel sourceChannel = new FileInputStream(source).getChannel();
             FileChannel destChannel = new FileOutputStream(dest).getChannel()) {
            
            long fileSize = sourceChannel.size();
            long position = 0;
            long chunkSize = 1024 * 1024;
            
            while (position < fileSize && !isCancelled) {
                long remaining = fileSize - position;
                long transferSize = Math.min(chunkSize, remaining);
                
                long transferred = destChannel.transferFrom(sourceChannel, position, transferSize);
                position += transferred;
                bytesCopied += transferred;
                
                int progress = (int) ((bytesCopied * 100) / totalBytesToCopy);
                updateProgress("Copying: " + source.getName(), progress);
            }
        }
    }

    private void copyWithStreams(File source, File dest) throws IOException {
        try (FileInputStream fis = new FileInputStream(source);
             FileOutputStream fos = new FileOutputStream(dest)) {
            
            byte[] buffer = new byte[1024 * 1024];
            int length;
            long fileBytesCopied = 0;
            long fileSize = source.length();
            
            while ((length = fis.read(buffer)) > 0 && !isCancelled) {
                fos.write(buffer, 0, length);
                fileBytesCopied += length;
                bytesCopied += length;
                
                int progress = (int) ((bytesCopied * 100) / totalBytesToCopy);
                updateProgress("Copying: " + source.getName(), progress);
            }
        }
    }

    private void copyLockedFile(File source, File dest) throws IOException {
        updateProgress("Copying (locked): " + source.getName(), 0);
        
        try (FileInputStream fis = new FileInputStream(source.getAbsolutePath());
             FileOutputStream fos = new FileOutputStream(dest.getAbsolutePath())) {
            
            byte[] buffer = new byte[1024 * 1024];
            int length;
            long fileBytesCopied = 0;
            long fileSize = source.length();
            
            while ((length = fis.read(buffer)) > 0 && !isCancelled) {
                fos.write(buffer, 0, length);
                fileBytesCopied += length;
                bytesCopied += length;
                
                int progress = (int) ((bytesCopied * 100) / totalBytesToCopy);
                updateProgress("Copying (locked): " + source.getName(), progress);
            }
        }
    }

    private void copyLockedFileWithForceClose(File source, File dest) throws IOException {
        // First try normal copy
        try {
            copyWithStreams(source, dest);
            return;
        } catch (IOException e) {
            log("File is locked, attempting to close locking process...");
        }
        
        // If normal copy failed, try to close the locking process
        try {
            String filePath = source.getAbsolutePath();
            if (closeLockingProcess(filePath)) {
                // Retry copy after closing process
                copyWithStreams(source, dest);
                log("Successfully copied after closing locking process");
            } else {
                throw new IOException("Failed to close process locking the file: " + filePath);
            }
        } catch (Exception e) {
            throw new IOException("Error handling locked file: " + e.getMessage());
        }
    }

    private boolean closeLockingProcess(String filePath) {
        // Try multiple locations for handle.exe
        String[] handleLocations = {
            "handle.exe",                         // Current directory
            "lib\\handle.exe",                    // lib subdirectory
            System.getenv("ProgramFiles") + "\\SysInternals\\handle.exe",
            System.getenv("ProgramFiles(x86)") + "\\SysInternals\\handle.exe",
            "C:\\Windows\\System32\\handle.exe"
        };

        for (String handlePath : handleLocations) {
            try {
                File handleExe = new File(handlePath);
                if (!handleExe.exists()) continue;

                // Build the command
                String[] cmd = {
                    "cmd", "/c",
                    handleExe.getAbsolutePath(),
                    filePath
                };

                // Run handle.exe to find locking process
                Process handleProcess = Runtime.getRuntime().exec(cmd);
                String output = readProcessOutput(handleProcess);

                // Parse the output to find PID
                Optional<String> pid = Arrays.stream(output.split("\\r?\\n"))
                    .filter(line -> line.contains(".exe"))
                    .findFirst()
                    .map(line -> line.split("pid:")[1].trim().split(" ")[0]);

                if (pid.isPresent()) {
                    // Terminate the process
                    Runtime.getRuntime().exec(new String[] {
                        "taskkill", "/F", "/PID", pid.get()
                    }).waitFor();
                    log("Successfully terminated process PID: " + pid.get());
                    return true;
                }
            } catch (Exception e) {
                log("Error with handle.exe at " + handlePath + ": " + e.getMessage());
            }
        }
        
        log("Could not find or use handle.exe to close locking process");
        return false;
    }

    private String readProcessOutput(Process process) throws IOException {
        StringBuilder output = new StringBuilder();
        try (BufferedReader reader = new BufferedReader(
            new InputStreamReader(process.getInputStream()))) {
            String line;
            while ((line = reader.readLine()) != null) {
                output.append(line).append("\n");
            }
        }
        return output.toString();
    }

    private void copyDirectory(File source, File dest) throws IOException {
        if (isCancelled) return;
        
        if (!dest.exists()) {
            dest.mkdir();
        }
        
        File[] files = source.listFiles();
        if (files != null) {
            for (File file : files) {
                if (isCancelled || shouldSkip(file) || isPriorityItem(file.getAbsolutePath())) continue;
                
                File destFile = new File(dest, file.getName());
                if (file.isDirectory()) {
                    copyDirectory(file, destFile);
                } else {
                    copySingleFile(file, destFile);
                }
            }
        }
    }

    private boolean isFileLocked(File file) {
        try {
            try (FileChannel channel = new FileOutputStream(file, true).getChannel()) {
                channel.tryLock();
            }
            return false;
        } catch (IOException e) {
            return true;
        }
    }

    private void updateProgress(String currentFile, int progress) {
        if (progressUpdater != null) {
            SwingUtilities.invokeLater(() -> {
                progressUpdater.updateProgress(currentFile, progress, bytesCopied, totalBytesToCopy);
            });
        }
    }

    private void log(String message) {
        if (progressUpdater != null) {
            SwingUtilities.invokeLater(() -> {
                progressUpdater.logMessage(message);
            });
        }
    }

    private String formatSize(long bytes) {
        if (bytes < 1024) return bytes + " B";
        if (bytes < 1024 * 1024) return String.format("%.2f KB", bytes / 1024.0);
        if (bytes < 1024 * 1024 * 1024) return String.format("%.2f MB", bytes / (1024.0 * 1024));
        return String.format("%.2f GB", bytes / (1024.0 * 1024 * 1024));
    }
}
