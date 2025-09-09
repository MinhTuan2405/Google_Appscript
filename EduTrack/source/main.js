function onOpen () {
  // get the current spread sheet app
  let ui = SpreadsheetApp.getUi ()
  
  // whole menu
  ui.createMenu ('EduTrack')
    .addItem ('About Our Products', 'Production_information') // ref: product_information.gs
    .addSeparator ()

    .addSubMenu (ui.createMenu ("Form") // ref: Form.gs
                  .addItem ('Form Creation', 'mainFormBuilder') 
                  .addItem ('Sync Result', 'manualSync')
                  .addItem ('Download class list', 'downloalClassList')) 
    .addSeparator ()

    .addSubMenu (ui.createMenu ("Folder") // ref: Folder.gs
                  .addItem('Open Folder Creator', 'showFolderCreatorSidebar')
                  .addItem ('Change Folder Structure', 'showChangeFolderStructureSidebar'))
    .addSeparator ()

    .addSubMenu (ui.createMenu ('Application') // ref: Application.gs
                  .addItem ('Clear Cache', 'clearCache')
                  .addItem ('Clear Sheet', 'clearSheet'))
    .addSeparator ()


    .addItem ('Update Permissions', 'updatePerrmissions')
    .addSeparator ()

    .addItem ('Apply rules', 'applyRules')
    .addSeparator ()


    // hihi
    // end create menu process
    .addToUi ()

}