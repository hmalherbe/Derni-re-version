function downloadAllFiles() {
  // Get the root folder of the script project
  var rootFolder = ScriptApp.getScriptProject().getRootFolder();
  
  // Iterate through all files in the root folder (and potentially subfolders)
  var files = rootFolder.getFiles();
  for (var i = 0; i < files.length; i++) {
    var file = files[i];
    
    // Download the file content
    var content = file.getBlob().getDataAsString();
    
    // (Optional) Save the content to a local file (requires Drive API)
     var filename = file.getName();
     var folderId = 'your_destination_folder_id'; // Replace with your folder ID
     DriveApp.createFile(filename, content, MimeType.PLAIN_TEXT).moveToFolder(folderId);
  }
}
