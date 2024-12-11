// getSelectedFileName.js
var shell = new ActiveXObject( "Shell.Application" );
var folder = shell.Windows().Item().Document; // Get the active File Explorer window

// Get the currently selected items
var selectedItems = folder.SelectedItems();

// Check if there are selected items
if ( selectedItems.Count > 0 ) {
  // Get the name of the first selected item
  var selectedFileName = selectedItems.Item().Name;
  WScript.Echo( "Selected File: " + selectedFileName ); // Display the file name in an alert
} else {
  WScript.Echo( "No file is selected." );
}