function onOpen () {
  // get the current spread sheet app
  let ui = SpreadsheetApp.getUi ()
  
  // whole menu
  ui.createMenu ('🛠️Tools')
    .addItem ('About Our Products', 'Production_information') // ref: product_information.gs
    .addSeparator ()

    .addSubMenu (ui.createMenu ("Form") // ref: Form.gs
                  .addItem ('Form Creation', 'mainFormBuilder') 
                  .addItem ('Sync Result', 'syncData')
                  .addItem ('Download class list', 'downloalClassList')) 
    .addSeparator ()

    .addSubMenu (ui.createMenu ("Folder")
                  .addItem ('Create folder with template', 'createFolderWithTemplate')
                  .addItem ('Create folder', 'createFolder'))
    .addSeparator ()

    // testing 
    // hihi
    // end create menu process
    .addToUi ()

}