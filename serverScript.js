function doGet(request) {
    // initiate html service
    return HtmlService.createTemplateFromFile('Index')
        .evaluate();
  }
  function include(filename) {
    // include css and js into html template (printing scriptlets)
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
  }
  
  function retrieveFolders(){
    // retrieve all folders and display names
    let folders = DriveApp.getFolders();
    while (folders.hasNext()) {
      let folder = folders.next();
      Logger.log("Folder: %s, id %s",folder.getName(), folder.getId());
    }
  }
  
  function getFoldersInFolder(folderId) {
    // retrieve the given folder, otherwise root folder is used
    // return a list of folder in JSON: {[{id, name}, ...]}
  
    // get folder
    let folder;
    if (folderId) {
      folder = DriveApp.getFolderById(folderId);
    } else {
      folder = DriveApp.getRootFolder();
    }
    
    // get subfolders
    let folders = folder.getFolders();
    let folderList = [];
    while (folders.hasNext()) {
      let subFolder = folders.next();
      folderList.push({
        id: subFolder.getId(),
        name: subFolder.getName()
      });
    }
    return JSON.stringify(folderList);
  }
  
  function getFolderPath(folderId) {
    // get a folder's path structure using its id
    // return a list folder path in JSON {[{id, name}, ...]}
    // the first element should be the root folder "Drive App", the last element should be the 
    // given folder if it's not root
  
    // get folder
    let path = [];
    let folder;
    if (folderId) {
      folder = DriveApp.getFolderById(folderId);
    }
    else {
      folder = DriveApp.getRootFolder();
    }
    
    // get all parent folders
    while (folder) {
      path.unshift({
        id: folder.getId(),
        name: folder.getName()
      });
  
      let parents = folder.getParents();
      if (parents.hasNext()) {
        folder = parents.next();
      }
      else {
        folder = null;
      }
    }
    return JSON.stringify(path);
    
  }
  
  function getFolderPathString(folderId){
    // construct a path string "Drive App/.../.../folder"
    return JSON.parse(getFolderPath(folderId)).map(folder => folder.name).join("/");
  }

  function processSheet(sheetLink){
    const validSheetLink = sheetLink? sheetLink : 'Placeholder';
    res = {
        'status': 'ok',
        'message': 'Sheet processed successfully',
        'sheetLink': validSheetLink
    }
    Logger.log("Response: %s", JSON.stringify(res));
    return JSON.stringify(res)
    
  }